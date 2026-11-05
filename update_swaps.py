import base64
import datetime
import glob
import json
import os
import pathlib
import re
import time
import pandas as pd

RBNZ_PAGE = "https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/wholesale-interest-rates"
CURR_FALLBACK = "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b2/hb2-daily-close.xlsx"
HIST_FALLBACK = "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b2/hb2-daily-close-1985-2017.xlsx"

SERIES_MAP = {
    "ocr": "INM.DP1.N",
    "b30": "INM.DB01.NZZV",
    "b60": "INM.DB02.NZZV",
    "b90": "INM.DB03.NZZV",
    "s1": "INM.DS01.NZZC",
    "s2": "INM.DS02.NZZC",
    "s3": "INM.DS03.NZZC",
    "s4": "INM.DS04.NZZC",
    "s5": "INM.DS05.NZZC",
    "s7": "INM.DS07.NZZC",
    "s10": "INM.DS10.NZZC",
    "s15": "INM.DS15.NZZC",
    "sp2_10": "INM.DS61.NZZC",
}


def is_valid_xlsx(data: bytes) -> bool:
    return bool(data and len(data) > 10000 and data[:4] == b"PK\x03\x04")


def add_cache_buster(url: str) -> str:
    sep = "&" if "?" in url else "?"
    return f"{url}{sep}_cb={int(time.time())}"


def try_curl_cffi_download(need_historical: bool, curr_dest: str, hist_dest: str) -> bool:
    """Tier 1: Chrome TLS (JA3/JA4) + HTTP/2 impersonation via curl_cffi."""
    try:
        from curl_cffi import requests as cf_requests
    except ImportError:
        print("[Tier 1 - curl_cffi] Package not installed; skipping.")
        return False

    print("[Tier 1 - curl_cffi] Attempting Chrome TLS impersonation fetch...")
    try:
        session = cf_requests.Session(impersonate="chrome")
        headers = {
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Referer": RBNZ_PAGE,
        }
        curr_url, hist_url = CURR_FALLBACK, HIST_FALLBACK

        page_resp = session.get(add_cache_buster(RBNZ_PAGE), headers=headers, timeout=30)
        print(f"[Tier 1 - curl_cffi] RBNZ page HTTP {page_resp.status_code} ({len(page_resp.content):,} bytes)")
        if page_resp.status_code == 200:
            for href in re.findall(r'href=["\']([^"\']*\.xlsx[^"\']*)["\']', page_resp.text, flags=re.I):
                full = "https://www.rbnz.govt.nz" + href if href.startswith("/") else href
                low = full.lower()
                if "hb2-daily-close" in low and "1985" not in low:
                    curr_url = full
                elif "1985-2017" in low:
                    hist_url = full

        r_curr = session.get(
            add_cache_buster(curr_url),
            headers={**headers, "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*"},
            timeout=45,
        )
        print(f"[Tier 1 - curl_cffi] Current XLSX HTTP {r_curr.status_code} ({len(r_curr.content):,} bytes)")
        if not is_valid_xlsx(r_curr.content):
            snippet = r_curr.content[:160].decode("utf-8", errors="replace").replace("\n", " ")
            print(f"[Tier 1 - curl_cffi] Non-XLSX payload: {snippet}")
            return False

        pathlib.Path(curr_dest).write_bytes(r_curr.content)
        if need_historical:
            r_hist = session.get(add_cache_buster(hist_url), headers=headers, timeout=45)
            if not is_valid_xlsx(r_hist.content):
                return False
            pathlib.Path(hist_dest).write_bytes(r_hist.content)

        print("[Tier 1 - curl_cffi] Downloaded valid RBNZ XLSX.")
        return True
    except Exception as exc:
        print(f"[Tier 1 - curl_cffi] Error: {exc}")
        return False


def wait_for_cloudflare_clear(page, label: str, max_wait_sec: int = 35) -> None:
    """Poll until Cloudflare challenge clears and .xlsx links render in the DOM."""
    start = time.time()
    clicked_turnstile = False
    while time.time() - start < max_wait_sec:
        try:
            title = (page.title() or "").strip()
            xlsx_count = page.locator('a[href*=".xlsx"]').count()
            if xlsx_count > 0 and "just a moment" not in title.lower():
                print(f"[{label}] Page ready: title='{title}' | .xlsx links={xlsx_count}")
                return
            print(f"[{label}] Waiting ({int(time.time() - start)}s)... title='{title}' | .xlsx links={xlsx_count}")

            if not clicked_turnstile and (time.time() - start) > 8:
                for frame in page.frames:
                    if "challenges.cloudflare.com" in (frame.url or ""):
                        box = frame.locator("body").bounding_box()
                        if box:
                            page.mouse.click(box["x"] + 28, box["y"] + box["height"] / 2)
                            clicked_turnstile = True
                            print(f"[{label}] Clicked Cloudflare Turnstile checkbox.")
                            break
        except Exception as exc:
            print(f"[{label}] Poll note: {exc}")
        page.wait_for_timeout(2500)


def fetch_xlsx_in_browser(page, match_fn, fallback_url: str, dest_path: str, label: str) -> bool:
    """Download XLSX using the browser's own network/TLS stack and cf_clearance cookies."""
    target_url = fallback_url
    matched_locator = None

    for link in page.locator('a[href*=".xlsx"]').all():
        href = link.get_attribute("href") or ""
        if match_fn(href.lower()):
            target_url = "https://www.rbnz.govt.nz" + href if href.startswith("/") else href
            matched_locator = link
            break

    bust_url = add_cache_buster(target_url)
    print(f"[{label}] In-browser fetch: {bust_url}")

    try:
        res = page.evaluate(
            """async (url) => {
                const r = await fetch(url, { cache: 'no-store', credentials: 'include' });
                if (!r.ok) return { ok: false, status: r.status, b64: '' };
                const bytes = new Uint8Array(await r.arrayBuffer());
                let bin = '';
                for (let i = 0; i < bytes.length; i += 0x8000) {
                    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
                }
                return { ok: true, status: r.status, b64: btoa(bin) };
            }""",
            bust_url,
        )
        if res and res.get("ok") and res.get("b64"):
            raw = base64.b64decode(res["b64"])
            print(f"[{label}] In-browser fetch HTTP {res.get('status')} ({len(raw):,} bytes)")
            if is_valid_xlsx(raw):
                pathlib.Path(dest_path).write_bytes(raw)
                return True
        else:
            print(f"[{label}] In-browser fetch returned status={res.get('status') if res else 'None'}")
    except Exception as exc:
        print(f"[{label}] In-browser fetch note: {exc}")

    if matched_locator is not None:
        try:
            with page.expect_download(timeout=45000) as dl_info:
                matched_locator.click(force=True)
            dl_info.value.save_as(dest_path)
            raw = pathlib.Path(dest_path).read_bytes()
            print(f"[{label}] Click download saved ({len(raw):,} bytes)")
            if is_valid_xlsx(raw):
                return True
        except Exception as exc:
            print(f"[{label}] Click download note: {exc}")

    return False


def try_camoufox_download(need_historical: bool, curr_dest: str, hist_dest: str) -> bool:
    """Tier 2: Anti-detect Camoufox browser under Xvfb (zero CDP leak)."""
    try:
        from camoufox.sync_api import Camoufox
    except ImportError:
        print("[Tier 2 - Camoufox] Package not installed; skipping.")
        return False

    print("[Tier 2 - Camoufox] Launching anti-detect browser...")
    try:
        with Camoufox(headless=False, humanize=True) as browser:
            page = browser.new_page()
            page.goto(add_cache_buster(RBNZ_PAGE), wait_until="domcontentloaded", timeout=60000)
            wait_for_cloudflare_clear(page, "Tier 2 - Camoufox", max_wait_sec=35)

            ok_curr = fetch_xlsx_in_browser(
                page, lambda h: "hb2-daily-close" in h and "1985" not in h, CURR_FALLBACK, curr_dest, "Tier 2 - Camoufox"
            )
            ok_hist = True
            if need_historical and ok_curr:
                ok_hist = fetch_xlsx_in_browser(
                    page, lambda h: "1985-2017" in h, HIST_FALLBACK, hist_dest, "Tier 2 - Camoufox"
                )
            return ok_curr and ok_hist
    except Exception as exc:
        print(f"[Tier 2 - Camoufox] Error: {exc}")
        return False


def try_headed_chrome_download(need_historical: bool, curr_dest: str, hist_dest: str) -> bool:
    """Tier 3: Headed Google Chrome under Xvfb with native OS fingerprint (no fake UA mismatch)."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("[Tier 3 - Chrome] Playwright not installed; skipping.")
        return False

    print("[Tier 3 - Chrome] Launching headed Chrome...")
    profile_dir = pathlib.Path("/tmp/hlf_chrome_profile")
    profile_dir.mkdir(parents=True, exist_ok=True)
    try:
        with sync_playwright() as p:
            kw = dict(
                user_data_dir=str(profile_dir),
                headless=False,
                accept_downloads=True,
                viewport={"width": 1366, "height": 900},
                args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
            )
            try:
                context = p.chromium.launch_persistent_context(channel="chrome", **kw)
            except Exception:
                context = p.chromium.launch_persistent_context(**kw)

            page = context.pages[0] if context.pages else context.new_page()
            page.goto(add_cache_buster(RBNZ_PAGE), wait_until="domcontentloaded", timeout=60000)
            wait_for_cloudflare_clear(page, "Tier 3 - Chrome", max_wait_sec=35)

            ok_curr = fetch_xlsx_in_browser(
                page, lambda h: "hb2-daily-close" in h and "1985" not in h, CURR_FALLBACK, curr_dest, "Tier 3 - Chrome"
            )
            ok_hist = True
            if need_historical and ok_curr:
                ok_hist = fetch_xlsx_in_browser(
                    page, lambda h: "1985-2017" in h, HIST_FALLBACK, hist_dest, "Tier 3 - Chrome"
                )
            context.close()
            return ok_curr and ok_hist
    except Exception as exc:
        print(f"[Tier 3 - Chrome] Error: {exc}")
        return False


def download_rbnz_files(need_historical: bool, curr_dest: str, hist_dest: str) -> bool:
    if try_curl_cffi_download(need_historical, curr_dest, hist_dest):
        return True
    if try_camoufox_download(need_historical, curr_dest, hist_dest):
        return True
    if try_headed_chrome_download(need_historical, curr_dest, hist_dest):
        return True
    return False


def load_existing_json(json_path: pathlib.Path) -> pd.DataFrame:
    raw = json.loads(json_path.read_text(encoding="utf-8"))
    data = {"date": pd.to_datetime(raw["d"])}
    for key in SERIES_MAP:
        scale = 1.0 if key == "sp2_10" else 100.0
        vals = raw.get(key, [None] * len(raw["d"]))
        data[key] = [None if v is None else float(v) / scale for v in vals]
    return pd.DataFrame(data)


def parse_rbnz_excel(filepath: str) -> pd.DataFrame:
    xls = pd.ExcelFile(filepath)
    _ = xls.sheet_names
    raw = pd.read_excel(xls, sheet_name="Data", header=None)
    sids = raw.iloc[4].astype(str).str.strip().tolist()
    df = raw.iloc[5:].copy()
    df.columns = ["date"] + sids[1:]
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"])

    out = {"date": df["date"]}
    for key, sid in SERIES_MAP.items():
        if sid in df.columns:
            out[key] = pd.to_numeric(df[sid], errors="coerce")
        else:
            out[key] = pd.Series([None] * len(df), index=df.index, dtype="float64")
    return pd.DataFrame(out)


def to_int(val, scale=100):
    return None if pd.isna(val) else int(round(float(val) * scale))


def build_static_html(swaps: pd.DataFrame, latest_str: str, now_utc: str) -> str:
    last = swaps.iloc[-1]
    prev = swaps.iloc[-2] if len(swaps) > 1 else last

    def row_days_ago(days: int):
        cutoff = last["date"] - pd.Timedelta(days=days)
        sub = swaps[swaps["date"] <= cutoff]
        return sub.iloc[-1] if len(sub) > 0 else swaps.iloc[0]

    w1, m1, y1 = row_days_ago(7), row_days_ago(30), row_days_ago(365)

    bench_rows = [
        ("Official Cash Rate (OCR)", "INM.DP1.N", "ocr", False),
        ("30-Day Bank Bill", "INM.DB01.NZZV", "b30", False),
        ("60-Day Bank Bill", "INM.DB02.NZZV", "b60", False),
        ("90-Day Bank Bill (BKBM)", "INM.DB03.NZZV", "b90", False),
        ("1-Year Swap", "INM.DS01.NZZC", "s1", False),
        ("2-Year Swap", "INM.DS02.NZZC", "s2", False),
        ("3-Year Swap", "INM.DS03.NZZC", "s3", False),
        ("4-Year Swap", "INM.DS04.NZZC", "s4", False),
        ("5-Year Swap", "INM.DS05.NZZC", "s5", False),
        ("7-Year Swap", "INM.DS07.NZZC", "s7", False),
        ("10-Year Swap", "INM.DS10.NZZC", "s10", False),
        ("15-Year Swap", "INM.DS15.NZZC", "s15", False),
        ("2Y–10Y Swap Spread", "INM.DS61.NZZC", "sp2_10", True),
    ]

    def fmt_val(v, is_bps=False):
        if pd.isna(v):
            return "—"
        return f"{int(round(v)):+d} bps" if is_bps else f"{v:.2f}%"

    def fmt_chg(cur, old, is_bps=False):
        if pd.isna(cur) or pd.isna(old):
            return "—"
        diff = int(round(cur - old)) if is_bps else int(round((cur - old) * 100))
        return f"{diff:+d} bps"

    snap_tr = [
        f"<tr><td><strong>{lbl}</strong></td><td><code>{sid}</code></td>"
        f"<td><strong>{fmt_val(last[col], bps)}</strong></td>"
        f"<td>{fmt_chg(last[col], prev[col], bps)}</td>"
        f"<td>{fmt_chg(last[col], w1[col], bps)}</td>"
        f"<td>{fmt_chg(last[col], m1[col], bps)}</td>"
        f"<td>{fmt_chg(last[col], y1[col], bps)}</td></tr>"
        for lbl, sid, col, bps in bench_rows
    ]

    yr_df = swaps.copy()
    yr_df["year"] = yr_df["date"].dt.year
    ann_tr = []
    for yr, grp in yr_df.groupby("year"):
        end_r = grp.iloc[-1]
        ann_tr.append(
            f"<tr><td><strong>{yr}</strong></td><td>{fmt_val(end_r['ocr'])}</td>"
            f"<td>{fmt_val(end_r['s1'])}</td>"
            f"<td>{fmt_val(end_r['s2'])} ({grp['s2'].min():.2f}%–{grp['s2'].max():.2f}%)</td>"
            f"<td>{fmt_val(end_r['s5'])} ({grp['s5'].min():.2f}%–{grp['s5'].max():.2f}%)</td></tr>"
        )
    ann_tr.reverse()

    daily_tr = [
        f"<tr><td>{r['date'].strftime('%Y-%m-%d')}</td><td>{fmt_val(r['ocr'])}</td>"
        f"<td>{fmt_val(r['b90'])}</td><td>{fmt_val(r['s1'])}</td><td>{fmt_val(r['s2'])}</td>"
        f"<td>{fmt_val(r['s3'])}</td><td>{fmt_val(r['s4'])}</td><td>{fmt_val(r['s5'])}</td>"
        f"<td>{fmt_val(r['s7'])}</td><td>{fmt_val(r['s10'])}</td>"
        f"<td>{fmt_val(r['sp2_10'], True)}</td></tr>"
        for _, r in swaps.tail(60).iloc[::-1].iterrows()
    ]

    return f"""<!DOCTYPE html>
<html lang="en-NZ">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="robots" content="index, follow">
  <title>NZ Wholesale Swap Rates — Static Daily Data Tables (RBNZ Table B2) | Home Loan Factory</title>
  <meta name="description" content="Daily updated static HTML reference tables of New Zealand wholesale swap rates (1 to 15 years), bank bill yields, and Official Cash Rate compiled from Reserve Bank of New Zealand Table B2.">
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; color: #1e293b; background: #f4f8fa; margin: 0; padding: 24px 14px; line-height: 1.55; }}
    main {{ max-width: 1080px; margin: 0 auto; background: #ffffff; border-top: 6px solid #0054a6; border-radius: 12px; padding: 28px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }}
    h1 {{ margin: 0 0 8px; color: #0f172a; font-size: 1.55rem; letter-spacing: -0.02em; }}
    h2 {{ margin: 28px 0 10px; color: #0b2559; font-size: 1.12rem; }}
    p {{ color: #475569; font-size: 0.92rem; margin: 0 0 12px; }}
    a {{ color: #0054a6; font-weight: 600; }}
    .cta {{ display: inline-block; margin: 6px 8px 14px 0; padding: 9px 16px; background: #0054a6; color: #ffffff !important; text-decoration: none; border-radius: 999px; font-size: 0.86rem; }}
    .cta-sec {{ background: #eff6ff; color: #0054a6 !important; border: 1px solid #bfdbfe; }}
    .wrap {{ overflow-x: auto; margin-bottom: 18px; border: 1px solid #e2e8f0; border-radius: 8px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.84rem; font-variant-numeric: tabular-nums; }}
    th {{ background: #f8fafc; color: #0b2559; text-align: left; padding: 10px 12px; border-bottom: 2px solid #cbd5e1; white-space: nowrap; }}
    td {{ padding: 8px 12px; border-bottom: 1px solid #e2e8f0; white-space: nowrap; }}
    tr:last-child td {{ border-bottom: none; }}
    code {{ background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-size: 0.78rem; color: #0f172a; }}
    .foot {{ margin-top: 22px; padding-top: 16px; border-top: 1px solid #e2e8f0; font-size: 0.82rem; color: #64748b; }}
  </style>
</head>
<body>
  <main>
    <h1>New Zealand Wholesale Swap Rates — Daily Reference Tables</h1>
    <p>Machine-readable static HTML tables updated automatically each New Zealand business day from <a href="https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/wholesale-interest-rates">Reserve Bank of New Zealand (RBNZ) Table B2</a> (sourced via RBNZ, Reuters, and NZFMA; 1-business-day delayed end-of-day closing mid-rates). Latest RBNZ close: <strong>{latest_str}</strong> (Checked: <code>{now_utc}</code>).</p>
    <div>
      <a class="cta" href="https://www.homeloanfactory.co.nz/tools/nz-wholesale-swap-rates">&larr; Open Interactive Swap Rates Chart</a>
      <a class="cta cta-sec" href="https://www.homeloanfactory.co.nz/tools/compare-nz-mortgage-rates">Compare Retail NZ Mortgage Rates</a>
    </div>

    <h2>1. Latest Wholesale Swap &amp; Bank Bill Closes ({latest_str})</h2>
    <div class="wrap">
      <table>
        <thead><tr><th>Benchmark</th><th>RBNZ Series ID</th><th>Latest Close</th><th>1D Move</th><th>1W Move</th><th>1M Move</th><th>1Y Move</th></tr></thead>
        <tbody>{''.join(snap_tr)}</tbody>
      </table>
    </div>

    <h2>2. Annual Wholesale Swap Summary (2011–Present)</h2>
    <div class="wrap">
      <table>
        <thead><tr><th>Year</th><th>Year-End OCR</th><th>Year-End 1Y Swap</th><th>Year-End 2Y Swap (Annual Range)</th><th>Year-End 5Y Swap (Annual Range)</th></tr></thead>
        <tbody>{''.join(ann_tr)}</tbody>
      </table>
    </div>

    <h2>3. Recent 60 Trading Days — Daily Closing Mid-Rates</h2>
    <div class="wrap">
      <table>
        <thead><tr><th>Date</th><th>OCR</th><th>90d Bill</th><th>1Y Swap</th><th>2Y Swap</th><th>3Y Swap</th><th>4Y Swap</th><th>5Y Swap</th><th>7Y Swap</th><th>10Y Swap</th><th>2Y–10Y Spread</th></tr></thead>
        <tbody>{''.join(daily_tr)}</tbody>
      </table>
    </div>

    <div class="foot">
      <p><strong>Data Attribution &amp; General Educational Disclosure:</strong> Source: Reserve Bank of New Zealand (Table B2), Reuters, and the New Zealand Financial Markets Association (NZFMA). Wholesale swap rates are institutional treasury benchmarks and are not retail mortgage rates available to borrowers. Provided for general educational purposes only by Home Loan Factory (<a href="https://www.homeloanfactory.co.nz/public-disclosures">Public Disclosures</a> | <a href="https://www.homeloanfactory.co.nz/how-we-get-paid">How We Get Paid</a>).</p>
    </div>
  </main>
</body>
</html>"""


def main() -> None:
    root = pathlib.Path(__file__).resolve().parent
    json_path = root / "nz-swap-rates.json"
    html_path = root / "index.html"
    tmp_curr = root / "tmp_hb2_curr.xlsx"
    tmp_hist = root / "tmp_hb2_hist.xlsx"

    frames = []
    old_latest = ""
    old_count = 0

    if json_path.exists():
        existing_df = load_existing_json(json_path)
        old_latest = existing_df["date"].max().strftime("%Y-%m-%d")
        old_count = len(existing_df)
        frames.append(existing_df)
        need_hist = False
    else:
        need_hist = True

    # Always ingest any manually uploaded .xlsx files in repo root as well
    local_xlsx = sorted(glob.glob(str(root / "*.xlsx")))
    local_newer = False
    for lf in local_xlsx:
        if not os.path.basename(lf).startswith("tmp_"):
            ldf = parse_rbnz_excel(lf)
            lmax = ldf["date"].max().strftime("%Y-%m-%d")
            print(f"Loaded repo XLSX '{os.path.basename(lf)}' (latest={lmax}, rows={len(ldf)})")
            if lmax > old_latest:
                local_newer = True
            frames.append(ldf)
            if "1985" in os.path.basename(lf):
                need_hist = False

    try:
        fetched_ok = download_rbnz_files(need_hist, str(tmp_curr), str(tmp_hist))
    except Exception as exc:
        print(f"RBNZ live fetch exception: {exc}")
        fetched_ok = False

    for tmp_file in (tmp_hist, tmp_curr):
        if tmp_file.exists():
            tdf = parse_rbnz_excel(str(tmp_file))
            print(f"Parsed live RBNZ '{tmp_file.name}': latest trading date = {tdf['date'].max().strftime('%Y-%m-%d')} ({len(tdf)} rows)")
            frames.append(tdf)
            os.remove(tmp_file)

    if not fetched_ok and not local_newer:
        raise RuntimeError(
            "All 3 live RBNZ download tiers were blocked by Cloudflare (and no newer .xlsx was uploaded to the repo). "
            "Failing workflow so GitHub Actions reports the block honestly instead of faking a green run."
        )

    if not frames:
        raise RuntimeError("No RBNZ data sources found to build nz-swap-rates.json.")

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.sort_values("date").drop_duplicates(subset=["date"], keep="last")
    swaps = combined[combined["date"] >= "2011-01-05"].dropna(subset=["s2"]).reset_index(drop=True)

    if len(swaps) < max(3900, old_count):
        raise RuntimeError(f"Unexpectedly low row count ({len(swaps)}); aborting overwrite.")

    latest_str = swaps["date"].iloc[-1].strftime("%Y-%m-%d")
    now_utc = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    payload = {
        "source": "Reserve Bank of New Zealand (Table B2), Reuters, NZFMA",
        "latest_date": latest_str,
        "generated_utc": now_utc,
        "count": len(swaps),
        "d": swaps["date"].dt.strftime("%Y-%m-%d").tolist(),
    }
    for key in SERIES_MAP:
        scale = 1 if key == "sp2_10" else 100
        payload[key] = [to_int(x, scale) for x in swaps[key]]

    json_path.write_text(json.dumps(payload, separators=(",", ":")), encoding="utf-8")
    html_path.write_text(build_static_html(swaps, latest_str, now_utc), encoding="utf-8")
    print(f"Saved nz-swap-rates.json and index.html through {latest_str} ({len(swaps)} rows, checked {now_utc}).")


if __name__ == "__main__":
    main()
