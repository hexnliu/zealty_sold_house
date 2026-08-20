# Zealty Sold House Scraper

A Python scraper that collects sold detached-house listing information from **Zealty.ca** and exports the results to an Excel workbook.

The scraper uses **Playwright** for browser automation, supports authenticated Zealty sessions, automatically processes multiple result pages, collects property-level details, and saves checkpoints so interrupted jobs can resume.

> **Important:** This project is intended for personal/research use. Make sure your use of the data complies with Zealty's Terms of Use and any applicable MLS®, CREA, or real-estate board data restrictions.

## Features

* Scrapes sold detached-house listings from Zealty
* Processes all configured search-result pages
* Collects individual property URLs
* Opens each property detail page
* Supports authenticated Zealty data
* Automatically logs in when the saved session expires
* Stores authentication state for future runs
* Supports headless execution for servers and scheduled jobs
* Saves progress after each property
* Resumes interrupted scraping jobs
* Retries previously failed properties
* Deduplicates property URLs
* Exports results to Excel
* Stores complex history data as JSON where appropriate
* Keeps credentials outside the source code using environment variables

## Requirements

* Python 3.10+
* Chromium
* A Zealty account
* Internet connection

Python packages:

```text
playwright
beautifulsoup4
pandas
openpyxl
```

## Installation

Clone the repository:

```bash
git clone <your-repository-url>
cd <your-repository-name>
```

Install the required Python packages:

```bash
pip install beautifulsoup4 playwright pandas openpyxl
```

Install Playwright Chromium:

```bash
playwright install chromium
```

For Python 3.13 on Windows, you can use:

```bat
py -3.13 -m pip install beautifulsoup4 playwright pandas openpyxl
py -3.13 -m playwright install chromium
```

## Zealty Credentials

Do **not** put your Zealty email address or password directly into the Python source code.

The scraper reads credentials from these environment variables:

```text
ZEALTY_EMAIL
ZEALTY_PASSWORD
```

### Windows

Open Command Prompt and run:

```bat
setx ZEALTY_EMAIL "your-email@example.com"
setx ZEALTY_PASSWORD "your-password"
```

Close and reopen Command Prompt after running `setx`.

You can verify that the email variable exists with:

```bat
echo %ZEALTY_EMAIL%
```

Avoid printing the password variable.

### Linux

```bash
export ZEALTY_EMAIL="your-email@example.com"
export ZEALTY_PASSWORD="your-password"
```

For a production server, configure these variables using your hosting platform, container, CI/CD system, or secret-management service.

## Running the Scraper

Run:

```bash
python zealty_scraper.py
```

On Windows with Python 3.13:

```bat
py -3.13 zealty_scraper.py
```

The scraper will:

1. Start Chromium.
2. Load an existing Zealty authentication session if available.
3. Check whether the session is still authenticated.
4. Automatically sign in when necessary.
5. Save the refreshed authentication state.
6. Visit each configured search-results page.
7. Collect unique property URLs.
8. Visit each property detail page.
9. Extract the available property information.
10. Save progress continuously.
11. Export the completed dataset to Excel.

## Search Configuration

The default search is based on:

```text
https://www.zealty.ca/region/bc/metro?propertyTypes=HSE&status=sold&sortBy=pricePerLotSize&month=30
```

It searches for:

* Property type: Detached house
* Status: Sold
* Region: Metro BC
* Period: Last 30 months
* Sorting: Price per lot size

Pagination follows this format:

```text
Page 1
/region/bc/metro?propertyTypes=HSE...

Page 2
/region/bc/metro/page/2?propertyTypes=HSE...

Page 3
/region/bc/metro/page/3?propertyTypes=HSE...
```

The number of pages can be configured in the Python script.

## Output Files

The scraper creates several files.

### `zealty_sold_houses.xlsx`

The final Excel workbook containing the collected property information.

Depending on the information available for each listing, fields may include:

* Address
* MLS number
* Listing URL
* Description
* Asking price
* Sold price
* Taxes
* Bedrooms
* Bathrooms
* Floor area
* Property age
* Lot/land size
* Frontage
* Zoning
* PID
* Parking
* Heating
* Water
* Sewer
* Listing dates
* Sold dates
* Market information
* Agent information
* Amenities
* Room information
* MLS pricing history
* Assessment history

Some complex sections may be stored as JSON text inside an Excel cell.

### `zealty_property_urls.csv`

Contains the unique property-detail URLs discovered from the search-result pages.

This is useful for troubleshooting and verifying how many properties were found before detailed scraping begins.

### `zealty_checkpoint.csv`

Stores intermediate scraping results.

The checkpoint allows a long-running scrape to recover from:

* Network failures
* Server restarts
* Script crashes
* Playwright errors
* Temporary Zealty errors

Successfully processed properties do not need to be scraped again when the script restarts.

Failed properties can be retried on a subsequent run.

### `zealty_auth.json`

Contains the Playwright authentication state used to maintain the Zealty login session.

**This file is sensitive.**

It may contain cookies and other authentication information. Never publish it or commit it to Git.

## Recommended `.gitignore`

Create a `.gitignore` file containing:

```gitignore
# Zealty authentication
zealty_auth.json

# Scraper output
zealty_checkpoint.csv
zealty_property_urls.csv
zealty_sold_houses.xlsx

# Environment variables
.env
.env.*

# Python
__pycache__/
*.py[cod]
*$py.class

# Virtual environments
venv/
.venv/
env/

# IDE
.vscode/
.idea/

# OS files
.DS_Store
Thumbs.db
```

## Authentication

Zealty exposes additional property information to authenticated users.

The scraper therefore uses the following authentication strategy:

```text
Start
  |
  v
Does zealty_auth.json exist?
  |
  +---- Yes ----> Load saved session
  |                   |
  |                   v
  |             Is session valid?
  |                /       \
  |              Yes        No
  |               |          |
  |               |          v
  |               |     Login again
  |               |          |
  |               |          v
  |               |     Save new session
  |               |          |
  +---------------+----------+
                  |
                  v
             Start scraping
```

This reduces unnecessary login requests.

If the existing session expires, the scraper attempts to authenticate again using:

```text
ZEALTY_EMAIL
ZEALTY_PASSWORD
```

After successful authentication, the new browser state is saved to:

```text
zealty_auth.json
```

## Headless Server Operation

The scraper is designed so it can eventually run without an interactive desktop.

For unattended operation, Chromium should run with:

```python
headless=True
```

The server must have the environment variables configured:

```text
ZEALTY_EMAIL
ZEALTY_PASSWORD
```

and Playwright Chromium must be installed:

```bash
playwright install chromium
```

On some Linux servers, Playwright dependencies may also be required:

```bash
playwright install --with-deps chromium
```

## Scheduled Jobs

### Linux Cron Example

Open the crontab:

```bash
crontab -e
```

For example, to run every day at 3:00 AM:

```cron
0 3 * * * cd /path/to/zealty-scraper && /usr/bin/python3 zealty_scraper.py >> scraper.log 2>&1
```

Make sure the scheduled environment has access to the required credentials.

### Windows Task Scheduler

The scraper can also be run using Windows Task Scheduler.

Program:

```text
C:\Users\<username>\AppData\Local\Programs\Python\Python313\python.exe
```

Arguments:

```text
C:\path\to\zealty_scraper.py
```

Set **Start in** to the project's directory so authentication, checkpoint, and output files are written to the expected location.

## Resume Support

Property scraping can take a significant amount of time.

For this reason, the scraper continuously writes results to:

```text
zealty_checkpoint.csv
```

If the process stops unexpectedly, simply run:

```bash
python zealty_scraper.py
```

again.

The scraper can use the checkpoint to avoid unnecessarily repeating successfully completed properties.

Do not delete the checkpoint file during an incomplete run unless you intentionally want to start over.

## Starting From Scratch

To perform a completely fresh scrape, stop the scraper and remove:

```text
zealty_checkpoint.csv
zealty_property_urls.csv
zealty_sold_houses.xlsx
```

Normally, keep:

```text
zealty_auth.json
```

so the existing authenticated session can be reused.

To force a completely new login as well, delete:

```text
zealty_auth.json
```

The next run will attempt to authenticate again using the configured environment variables.

## Troubleshooting

### `ModuleNotFoundError: No module named 'bs4'`

Install BeautifulSoup:

```bash
pip install beautifulsoup4
```

For Python 3.13 on Windows:

```bat
py -3.13 -m pip install beautifulsoup4
```

### Playwright Chromium Is Missing

If you receive a message indicating that the browser executable does not exist:

```bash
playwright install chromium
```

or:

```bat
py -3.13 -m playwright install chromium
```

### `Event loop is closed! Is Playwright already stopped?`

All Playwright operations must execute while:

```python
with sync_playwright() as p:
```

is still active.

Do not call `page.goto()`, `context.close()`, or other Playwright methods after leaving that block.

Also avoid manually calling:

```python
p.stop()
```

when using the `with sync_playwright()` context manager.

### Search Results Work but Property Details Are Missing

This usually indicates that the browser is not authenticated.

Check:

1. `ZEALTY_EMAIL` is configured.
2. `ZEALTY_PASSWORD` is configured.
3. Login succeeds.
4. `zealty_auth.json` is created.
5. The authenticated session has not expired.
6. Zealty has not introduced an additional verification step.

Deleting `zealty_auth.json` forces the scraper to create a new authenticated session on its next run.

### Automated Login Fails

Automated authentication can fail if Zealty changes:

* Login page HTML
* Input names or labels
* Button text
* Authentication flow
* Session handling

It can also fail if additional verification is introduced, such as:

* CAPTCHA
* Email verification
* Multi-factor authentication
* Bot protection

If that happens, the login selectors or authentication strategy may need to be updated.

## Security

Never commit any of the following to a public repository:

```text
Zealty password
zealty_auth.json
.env
authentication cookies
session tokens
```

Use environment variables or a secret-management service for production deployments.

If an authentication file or password is accidentally committed to GitHub, removing it in a later commit may not be sufficient because it can remain in Git history. Revoke/rotate the affected credentials and remove the secret from the repository history.

## Responsible Use

This project automates access to information available through a Zealty account.

Before using the scraper, review the applicable:

* Zealty Terms of Use
* Website access restrictions
* MLS® data restrictions
* CREA requirements
* Local real-estate board requirements
* Applicable privacy and data-protection requirements

Do not use the scraper to bypass CAPTCHA, access controls, account restrictions, rate limits, or other technical protections.

Use reasonable request intervals and avoid creating unnecessary load on the service.

## Project Structure

A typical repository will look like:

```text
zealty-scraper/
├── zealty_scraper.py
├── README.md
├── .gitignore
│
├── zealty_auth.json          # generated, DO NOT COMMIT
├── zealty_checkpoint.csv     # generated
├── zealty_property_urls.csv  # generated
└── zealty_sold_houses.xlsx   # generated
```

## Disclaimer

This project is not affiliated with, endorsed by, or sponsored by Zealty, CREA, MLS®, Greater Vancouver REALTORS®, or any other real-estate board or organization.

Website structure and authentication mechanisms can change at any time. As a result, the scraper may require maintenance when the source website changes.

Data collected by this project should be independently verified before being used for financial, investment, valuation, legal, or real-estate decisions.
