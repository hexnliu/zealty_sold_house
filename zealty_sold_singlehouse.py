import json
import re
import time
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

START_URL = "https://www.zealty.ca/region/bc/metro?propertyTypes=HSE&status=sold&sortBy=pricePerLotSize&month=30"
BASE = "https://www.zealty.ca"
OUTPUT_XLSX = Path("zealty_sold_houses.xlsx")
CHECKPOINT_CSV = Path("zealty_checkpoint.csv")
MAX_PAGES = 34
REQUEST_DELAY = 0.8


def clean_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def page_url(page_no: int) -> str:
    if page_no == 1:
        return START_URL
    return (
        f"{BASE}/region/bc/metro/page/{page_no}"
        "?propertyTypes=HSE&status=sold&sortBy=pricePerLotSize&month=30"
    )


def collect_listing_urls(page, max_pages=MAX_PAGES):
    urls = []
    seen = set()

    for page_no in range(1, max_pages + 1):
        url = page_url(page_no)
        print(f"[index {page_no}/{max_pages}] {url}")
        page.goto(url, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(800)

        hrefs = page.locator('a[href*="/property/"]').evaluate_all(
            "els => els.map(e => e.href)"
        )
        hrefs = [h.split("#")[0] for h in hrefs if "/property/" in h]

        new_count = 0
        for href in hrefs:
            if href not in seen:
                seen.add(href)
                urls.append(href)
                new_count += 1

        print(f"  found {len(hrefs)} property links, {new_count} new; total={len(urls)}")
        time.sleep(REQUEST_DELAY)

    return urls


def extract_table(table):
    rows = []
    for tr in table.select("tr"):
        cells = [clean_text(c.get_text(" ", strip=True)) for c in tr.select("th, td")]
        if cells:
            rows.append(cells)
    return rows


def parse_key_value_details(soup):
    """Extract property-detail label/value pairs using several DOM patterns."""
    out = {}

    # Standard tables / definition lists if present.
    for tr in soup.select("tr"):
        cells = [clean_text(x.get_text(" ", strip=True)) for x in tr.select(":scope > th, :scope > td")]
        if len(cells) == 2 and cells[0] and cells[1]:
            out.setdefault(cells[0], cells[1])

    for dt in soup.select("dt"):
        dd = dt.find_next_sibling("dd")
        if dd:
            k = clean_text(dt.get_text(" ", strip=True))
            v = clean_text(dd.get_text(" ", strip=True))
            if k and v:
                out.setdefault(k, v)

    # Zealty currently renders many detail rows as div-based label/value pairs.
    # This heuristic looks for compact containers whose text has exactly two
    # meaningful direct child chunks.
    for node in soup.select("div, li"):
        children = [
            clean_text(c.get_text(" ", strip=True))
            for c in node.find_all(recursive=False)
            if clean_text(c.get_text(" ", strip=True))
        ]
        if len(children) == 2:
            k, v = children
            if (
                1 <= len(k) <= 60
                and 1 <= len(v) <= 500
                and not k.startswith("$")
                and k.lower() not in {"property details", "description"}
            ):
                out.setdefault(k, v)

    return out


def section_text(soup, heading_text):
    heading = None
    for tag in soup.find_all(["h1", "h2", "h3", "h4"]):
        if clean_text(tag.get_text(" ", strip=True)).lower() == heading_text.lower():
            heading = tag
            break
    if not heading:
        return ""

    chunks = []
    for sib in heading.next_siblings:
        name = getattr(sib, "name", None)
        if name in {"h1", "h2", "h3", "h4"}:
            break
        if hasattr(sib, "get_text"):
            txt = clean_text(sib.get_text(" ", strip=True))
        else:
            txt = clean_text(sib)
        if txt:
            chunks.append(txt)
    return " ".join(chunks)


def extract_history_table(soup, heading_name):
    for h in soup.find_all(["h2", "h3", "h4"]):
        if heading_name.lower() in clean_text(h.get_text(" ", strip=True)).lower():
            table = h.find_next("table")
            if table:
                return extract_table(table)
    return []


def parse_property(html, url):
    soup = BeautifulSoup(html, "html.parser")
    row = {"URL": url}

    h1 = soup.find("h1")
    if h1:
        row["Address"] = clean_text(h1.get_text(" ", strip=True))

    full_text = clean_text(soup.get_text(" ", strip=True))

    m = re.search(r"MLS[®\s]*([A-Z]\d{6,8})", full_text, re.I)
    if m:
        row["MLS"] = m.group(1).upper()

    row["Description"] = section_text(soup, "Description")

    details = parse_key_value_details(soup)
    # Avoid overwriting the primary columns above.
    for k, v in details.items():
        if k not in row:
            row[k] = v

    # Capture structured multi-row sections as JSON so no detail is discarded.
    room_tables = []
    room_heading = next(
        (h for h in soup.find_all(["h2", "h3", "h4"])
         if "room information" in clean_text(h.get_text(" ", strip=True)).lower()),
        None,
    )
    if room_heading:
        for table in room_heading.find_all_next("table", limit=2):
            room_tables.append(extract_table(table))
    if room_tables:
        row["Room Information JSON"] = json.dumps(room_tables, ensure_ascii=False)

    pricing_hist = extract_history_table(soup, "MLS Pricing History")
    if pricing_hist:
        row["MLS Pricing History JSON"] = json.dumps(pricing_hist, ensure_ascii=False)

    assessment_hist = extract_history_table(soup, "Assessment History")
    if assessment_hist:
        row["Assessment History JSON"] = json.dumps(assessment_hist, ensure_ascii=False)

    # Last-resort raw Property Details text for fields that a future DOM change
    # may prevent the heuristic from splitting into columns.
    raw_details = section_text(soup, "Property Details")
    if raw_details:
        row["Property Details Raw"] = raw_details

    return row


def save_excel(rows, path=OUTPUT_XLSX):
    df = pd.DataFrame(rows)

    preferred = [
        "MLS", "Address", "Sold Price", "Final Asking Price", "Original Asking Price",
        "Property Taxes", "Floor Area", "Floor Area (Finished)", "Storeys (Finished)",
        "Age", "Land Size", "Land Frontage", "Land Dimensions", "Price per Land SqFt",
        "Zoning", "Listing Date", "Days on Market", "Sold Date", "Sale Reported Date",
        "Property Type", "Style of House", "Bedrooms", "Bathrooms", "Ownership Interest",
        "PID", "Seller's Agent", "Description", "URL",
    ]
    cols = [c for c in preferred if c in df.columns] + [c for c in df.columns if c not in preferred]
    df = df[cols]

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sold Houses")
        ws = writer.book["Sold Houses"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # Practical widths without making extremely long JSON columns enormous.
        for column_cells in ws.columns:
            letter = column_cells[0].column_letter
            max_len = max(len(str(c.value or "")) for c in column_cells[:200])
            ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 45)

    print(f"Saved {len(df)} rows to {path.resolve()}")


def main():
    rows = []
    completed = set()

    if CHECKPOINT_CSV.exists():
        old = pd.read_csv(CHECKPOINT_CSV, dtype=str).fillna("")
        if "URL" in old.columns:
            completed = set(old["URL"])
            rows = old.to_dict("records")
            print(f"Resuming from checkpoint: {len(completed)} properties already saved")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1440, "height": 1200},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/126.0 Safari/537.36"
            ),
        )
        page = context.new_page()

        urls = collect_listing_urls(page)
        print(f"Unique property URLs collected: {len(urls)}")

        for i, url in enumerate(urls, 1):
            if url in completed:
                continue

            print(f"[property {i}/{len(urls)}] {url}")
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(700)
                html = page.content()
                row = parse_property(html, url)
                rows.append(row)
                completed.add(url)

                # Checkpoint after every property.
                pd.DataFrame(rows).to_csv(CHECKPOINT_CSV, index=False, encoding="utf-8-sig")
            except Exception as exc:
                print(f"  ERROR: {exc}")
                rows.append({"URL": url, "Scrape Error": str(exc)})
                pd.DataFrame(rows).to_csv(CHECKPOINT_CSV, index=False, encoding="utf-8-sig")

            time.sleep(REQUEST_DELAY)

        browser.close()

    save_excel(rows)


if __name__ == "__main__":
    main()
