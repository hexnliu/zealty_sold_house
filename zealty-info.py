import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import time
import os
import re


BASE_URL = "https://www.zealty.ca/strata-buildings"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0 Safari/537.36"
    )
}

OUTPUT_DIR = "output"

# LISTING_FILE = os.path.join(
#     OUTPUT_DIR,
#     "bc_strata_buildings.xlsx"
# )

# DETAIL_FILE = os.path.join(
#     OUTPUT_DIR,
#     "bc_strata_buildings_details.csv"
# )

DETAIL_EXCEL = os.path.join(
    OUTPUT_DIR,
    "bc_strata_buildings_details.xlsx"
)


# ============================================================
# HTTP
# ============================================================

def create_session():

    session = requests.Session()

    session.headers.update(
        HEADERS
    )

    return session


def fetch_page(
    session,
    url,
    retries=3
):

    for attempt in range(
        1,
        retries + 1
    ):

        try:

            print(
                f"Fetching: {url}"
            )

            response = session.get(
                url,
                timeout=30
            )

            response.raise_for_status()

            return response.text

        except requests.RequestException as e:

            print(
                f"Request failed "
                f"({attempt}/{retries}): {e}"
            )

            if attempt < retries:

                time.sleep(
                    attempt * 2
                )

    return None


# ============================================================
# LISTING PAGE
# ============================================================

def get_page_url(page_number):

    if page_number == 1:

        return BASE_URL

    return (
        f"{BASE_URL}/page/{page_number}"
    )


def parse_listing_page(
    html,
    page_number,
    page_url
):

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    buildings = []

    seen_urls = set()

    for link in soup.find_all(
        "a",
        href=True
    ):

        text = link.get_text(
            " ",
            strip=True
        )

        if "Strata Plan:" not in text:

            continue

        building_url = urljoin(
            page_url,
            link["href"]
        )

        if building_url in seen_urls:

            continue

        seen_urls.add(
            building_url
        )

        record = {

            "BuildingURL":
                building_url,

            "PageNumber":
                page_number,

            "ListingText":
                text
        }

        # Extract strata plan

        match = re.search(
            r"Strata\s+Plan:\s*(.+)",
            text,
            re.IGNORECASE
        )

        if match:

            record["StrataPlan"] = (
                match.group(1).strip()
            )

        buildings.append(
            record
        )

    return buildings


# ============================================================
# DETAIL PAGE
# ============================================================

def clean_text(value):

    if value is None:

        return ""

    return " ".join(
        value.split()
    ).strip()


def add_field(
    record,
    name,
    value
):

    name = clean_text(
        name
    )

    value = clean_text(
        value
    )

    if not name or not value:

        return

    # Avoid overwriting an existing value
    if name not in record:

        record[name] = value

        return

    # If same field occurs more than once,
    # keep both values.
    existing = record[name]

    if value != existing:

        record[name] = (
            existing
            + " | "
            + value
        )


def parse_detail_page(
    html,
    building_url
):

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    record = {

        "BuildingURL":
            building_url
    }


    # ========================================================
    # 1. TITLE
    # ========================================================

    if soup.title:

        record["PageTitle"] = clean_text(
            soup.title.get_text()
        )


    # ========================================================
    # 2. META DESCRIPTION
    # ========================================================

    meta_description = soup.find(
        "meta",
        attrs={
            "name": "description"
        }
    )

    if meta_description:

        content = meta_description.get(
            "content"
        )

        if content:

            record["MetaDescription"] = (
                clean_text(content)
            )


    # ========================================================
    # 3. TABLE DATA
    # ========================================================

    for table in soup.find_all(
        "table"
    ):

        for row in table.find_all(
            "tr"
        ):

            cells = row.find_all(
                ["th", "td"]
            )

            if len(cells) < 2:

                continue

            values = [
                clean_text(
                    cell.get_text(
                        " ",
                        strip=True
                    )
                )
                for cell in cells
            ]

            # Handle:
            #
            # <th>Year Built</th>
            # <td>1995</td>
            #
            # or
            #
            # <td>Year Built</td>
            # <td>1995</td>

            for i in range(
                0,
                len(values) - 1
            ):

                label = values[i]

                value = values[i + 1]

                if not label or not value:

                    continue

                if (
                    len(label) < 100
                    and len(value) < 500
                ):

                    add_field(
                        record,
                        label,
                        value
                    )


    # ========================================================
    # 4. TD WITH TITLE
    # ========================================================

    for cell in soup.find_all(
        ["td", "div", "span"]
    ):

        title = cell.get(
            "title"
        )

        if not title:

            continue

        value = clean_text(
            cell.get_text(
                " ",
                strip=True
            )
        )

        if value:

            add_field(
                record,
                title,
                value
            )


    # ========================================================
    # 5. DT / DD
    # ========================================================

    for dt in soup.find_all(
        "dt"
    ):

        dd = dt.find_next_sibling(
            "dd"
        )

        if not dd:

            continue

        label = clean_text(
            dt.get_text(
                " ",
                strip=True
            )
        )

        value = clean_text(
            dd.get_text(
                " ",
                strip=True
            )
        )

        add_field(
            record,
            label,
            value
        )


    # ========================================================
    # 6. LABEL + VALUE
    # ========================================================

    for label in soup.find_all(
        "label"
    ):

        value = label.find_next_sibling()

        if not value:

            continue

        label_text = clean_text(
            label.get_text(
                " ",
                strip=True
            )
        )

        value_text = clean_text(
            value.get_text(
                " ",
                strip=True
            )
        )

        add_field(
            record,
            label_text,
            value_text
        )


    # ========================================================
    # 7. SEARCH FOR YEAR BUILT
    # ========================================================

    page_text = clean_text(
        soup.get_text(
            " ",
            strip=True
        )
    )

    year_patterns = [

        r"Year\s+Built\s*:?\s*(\d{4})",

        r"Built\s*:?\s*(\d{4})",

        r"Built\s+in\s+(\d{4})"
    ]

    for pattern in year_patterns:

        match = re.search(
            pattern,
            page_text,
            re.IGNORECASE
        )

        if match:

            record["YearBuilt"] = (
                match.group(1)
            )

            break


    # ========================================================
    # 8. PHOTO / IMAGE URLS
    # ========================================================

    photo_urls = []

    for img in soup.find_all(
        "img"
    ):

        possible_urls = [

            img.get("src"),

            img.get(
                "data-src"
            ),

            img.get(
                "data-lazy-src"
            ),

            img.get(
                "data-original"
            )
        ]

        for image_url in possible_urls:

            if not image_url:

                continue

            image_url = urljoin(
                building_url,
                image_url
            )

            if image_url not in photo_urls:

                photo_urls.append(
                    image_url
                )


    # ========================================================
    # 9. OPEN GRAPH IMAGE
    # ========================================================

    og_image = soup.find(
        "meta",
        property="og:image"
    )

    if og_image:

        image_url = og_image.get(
            "content"
        )

        if image_url:

            image_url = urljoin(
                building_url,
                image_url
            )

            if image_url not in photo_urls:

                photo_urls.append(
                    image_url
                )


    # ========================================================
    # 10. BACKGROUND IMAGES
    # ========================================================

    for tag in soup.find_all(
        style=True
    ):

        style = tag.get(
            "style"
        )

        matches = re.findall(
            r"url\(['\"]?([^'\")]+)",
            style,
            re.IGNORECASE
        )

        for image_url in matches:

            image_url = urljoin(
                building_url,
                image_url
            )

            if image_url not in photo_urls:

                photo_urls.append(
                    image_url
                )


    if photo_urls:

        record["PhotoURLs"] = "|".join(
            photo_urls
        )

        record["PhotoCount"] = len(
            photo_urls
        )

    else:

        record["PhotoURLs"] = ""

        record["PhotoCount"] = 0


    return record


# ============================================================
# SAVE
# ============================================================

def save_details(
    records
):

    if not records:

        return

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    df = pd.DataFrame(
        records
    )

    df = df.replace(
        r"^\s*$",
        pd.NA,
        regex=True
    )

    # Put important columns first

    first_columns = [

        "StrataPlan",

        "PageTitle",

        "YearBuilt",

        "BuildingURL",

        "PhotoCount",

        "PhotoURLs"
    ]

    first_columns = [

        col
        for col in first_columns
        if col in df.columns
    ]

    remaining_columns = [

        col
        for col in df.columns
        if col not in first_columns
    ]

    df = df[
        first_columns
        + remaining_columns
    ]

    # df.to_csv(
    #     DETAIL_FILE,
    #     index=False,
    #     encoding="utf-8-sig"
    # )

    df.to_excel(
        DETAIL_EXCEL,
        index=False
    )

    print(
        f"Saved detail records: "
        f"{len(df)}"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    # --------------------------------------------------------
    # SETTINGS
    # --------------------------------------------------------

    START_PAGE = 1

    # Test with 3 pages first
    END_PAGE = 764

    # Delay between listing pages
    LISTING_DELAY = 1.5

    # Delay between individual building pages
    DETAIL_DELAY = 1.5

    # --------------------------------------------------------

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    session = create_session()

    # ========================================================
    # STEP 1
    # SCRAPE LISTING PAGES
    # ========================================================

    all_buildings = []

    print()
    print(
        "========================================"
    )
    print(
        "STEP 1 - LISTING PAGES"
    )
    print(
        "========================================"
    )

    for page_number in range(
        START_PAGE,
        END_PAGE + 1
    ):

        url = get_page_url(
            page_number
        )

        html = fetch_page(
            session,
            url
        )

        if not html:

            continue

        buildings = parse_listing_page(
            html,
            page_number,
            url
        )

        print(
            f"Page {page_number}: "
            f"{len(buildings)} buildings"
        )

        all_buildings.extend(
            buildings
        )

        time.sleep(
            LISTING_DELAY
        )


    # Remove duplicate URLs

    listing_df = pd.DataFrame(
        all_buildings
    )

    listing_df = listing_df.drop_duplicates(
        subset=["BuildingURL"]
    )

    print()
    print(
        f"Total unique buildings: "
        f"{len(listing_df)}"
    )


    # # Save listing data

    # listing_df.to_excel(
    #     LISTING_EXCEL,
    #     index=False
    # )


    # ========================================================
    # STEP 2
    # FETCH INDIVIDUAL BUILDING PAGES
    # ========================================================

    print()
    print(
        "========================================"
    )
    print(
        "STEP 2 - BUILDING DETAILS"
    )
    print(
        "========================================"
    )


    details = []

    total = len(
        listing_df
    )


    for index, row in listing_df.iterrows():

        building_url = row[
            "BuildingURL"
        ]

        print()
        print(
            f"[{index + 1}/{total}]"
        )

        html = fetch_page(
            session,
            building_url
        )

        if not html:

            print(
                "Skipping"
            )

            continue


        try:

            detail = parse_detail_page(
                html,
                building_url
            )

            # Add listing information

            for column in listing_df.columns:

                if column not in detail:

                    detail[column] = row[
                        column
                    ]

            details.append(
                detail
            )

            print(
                "Year Built:",
                detail.get(
                    "YearBuilt",
                    ""
                )
            )

            print(
                "Photos:",
                detail.get(
                    "PhotoCount",
                    0
                )
            )


            # Save after EVERY building

            save_details(
                details
            )

        except Exception as e:

            print(
                "ERROR:",
                e
            )


        time.sleep(
            DETAIL_DELAY
        )


    print()
    print(
        "========================================"
    )

    print(
        "FINISHED"
    )

    print(
        f"Buildings processed: "
        f"{len(details)}"
    )

    # print(
    #     f"Output: {DETAIL_FILE}"
    # )

    print(
        f"Excel: {DETAIL_EXCEL}"
    )


if __name__ == "__main__":

    main()