import datetime
import glob
import json
import os
import pathlib
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


def fetch_xlsx_from_page(page, context, match_fn, fallback_url: str, dest_path: str) -> bool:
    target_url = fallback_url
    matched_locator = None

    for link in page.locator('a[href*=".xlsx"]').all():
        href = link.get_attribute("href") or ""
        if match_fn(href.lower()):
            target_url = "https://www.rbnz.govt.nz" + href if href.startswith("/") else href
            matched_locator = link
            break

    try:
        resp = context.request.get(target_url, timeout=60000)
        body = resp.body()
        if resp.ok and len(body) > 4000 and body[:2] == b"PK":
            pathlib.Path(dest_path).write_bytes(body)
            return True
    except Exception as exc:
        print(f"Direct context fetch note: {exc}")

    if matched_locator is not None:
        try:
            with page.expect_download(timeout=60000) as dl_info:
                matched_locator.click(force=True)
            dl_info.value.save_as(dest_path)
            if pathlib.Path(dest_path).read_bytes()[:2] == b"PK":
                return True
        except Exception as exc:
            print(f"Browser click download note: {exc}")

    return False


def download_rbnz_files(need_historical: bool, curr_dest: str, hist_dest: str) -> bool:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled"],
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1366, "height": 900},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        )
        page = context.new_page()
        page.goto(RBNZ_PAGE, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(5000)

        ok_curr = fetch_xlsx_from_page(
            page,
            context,
            lambda h: "hb2-daily-close" in h and "1985" not in h,
            CURR_FALLBACK,
            curr_dest,
        )
        ok_hist = True
        if need_historical:
            ok_hist = fetch_xlsx_from_page(
                page,
                context,
                lambda h: "1985-2017" in h,
                HIST_FALLBACK,
                hist_dest,
            )

        browser.close()
        return ok_curr and ok_hist


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
    <p>Machine-readable static HTML tables updated automatically each New Zealand business morning from <a href="https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/wholesale-interest-rates">Reserve Bank of New Zealand (RBNZ) Table B2</a> (sourced via RBNZ, Reuters, and NZFMA; 1-business-day delayed end-of-day closing mid-rates). Latest RBNZ close: <strong>{latest_str}</strong> (Checked: <code>{now_utc}</code>).</p>
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
        local_xlsx = sorted(glob.glob(str(root / "*.xlsx")))
        for lf in local_xlsx:
            if not os.path.basename(lf).startswith("tmp_"):
                frames.append(parse_rbnz_excel(lf))
        need_hist = not any("1985" in os.path.basename(lf) for lf in local_xlsx)

    try:
        fetched_ok = download_rbnz_files(need_hist, str(tmp_curr), str(tmp_hist))
    except Exception as exc:
        print(f"RBNZ browser fetch warning: {exc}")
        fetched_ok = False

    if not fetched_ok and (json_path.exists() or frames):
        print("Live RBNZ download skipped or unavailable; proceeding with stored/seed data.")

    for tmp_file in (tmp_hist, tmp_curr):
        if tmp_file.exists():
            frames.append(parse_rbnz_excel(str(tmp_file)))
            os.remove(tmp_file)

    if not frames:
        raise RuntimeError("No RBNZ data sources found to build nz-swap-rates.json.")

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.sort_values("date").drop_duplicates(subset=["date"], keep="last")
    swaps = combined[combined["date"] >= "2011-01-05"].dropna(subset=["s2"]).reset_index(drop=True)

    if len(swaps) < max(3900, old_count):
        raise RuntimeError(f"Unexpectedly low row count ({len(swaps)}); aborting overwrite.")

    latest_str = swaps["date"].iloc[-1].strftime("%Y-%m-%d")
    if latest_str == old_latest and len(swaps) == old_count and html_path.exists():
        print(f"Already up to date through {old_latest} ({old_count} rows).")
        return

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
    print(f"Saved nz-swap-rates.json and index.html through {latest_str} ({len(swaps)} rows).")


if __name__ == "__main__":
    main()
