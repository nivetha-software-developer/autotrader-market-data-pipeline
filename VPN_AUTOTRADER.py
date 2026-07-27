# ========================
# Standard Library Imports
# ========================
import glob
import random
import subprocess
import time
import copy
import traceback
from dotenv import load_dotenv
import numpy as np
from postmarker.core import PostmarkClient
import csv

# ========================
# Selenium (Web Automation)
# ========================
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ========================
# Other Imports
# ========================
import json
import os
from datetime import datetime,timedelta, date
import pandas as pd
from HelperFunctions import root_folder
from HelperFunctions import yesterdayDate, yesterdayDateInDash, todayDate, root_folder
from Utilities.DictList import DictList
from Utilities.Logger import Logger
import paramiko
import numpy as np

# Calculate yesterday and day-before-yesterday
today = date.today()
yesterday_output_file = (today - timedelta(days=1)).strftime("%Y%m%d")
day_before_yesterday_output_file = (today - timedelta(days=2)).strftime("%Y%m%d")

scrap_date = datetime.now().strftime('%Y-%m-%d')

env_path = os.path.join(root_folder, "SFTP", "env")
APPSETTING_FILE_PATH = os.path.join(root_folder, "Appsettings.json")
# ========================
# Runtime Configuration
# ========================

# Load application settings (parts window, labels, etc.)
with open(APPSETTING_FILE_PATH, "r") as f:
    settings = json.load(f)

start_make_with = settings["Part"]["StartsWith"]
end_make_with = settings["Part"]["EndsWith"]
part_label = settings["Part"]["PartLabel"]

# File that stores running progress for the current part/date (used to resume)
running_category_json_file = os.path.join(root_folder, "JSON", f"running_category_{part_label}_{todayDate}.json")
csv_filename = os.path.join(root_folder, "Output", f"autotrader_listings_{part_label}_{todayDate}.csv")

# Output file for unique car listings
output_filename = os.path.join(root_folder, "Output", f"autotrader_listings_output_{part_label}_{todayDate}.csv")

# Locate previous Excel files in Output folder
xlsx_files_path = os.path.join(root_folder, 'Output')
xlsx_files = glob.glob(os.path.join(xlsx_files_path, '*.csv'))

xlsx_files_path = os.path.join(root_folder, 'Output')
xlsx_history_files = os.path.join(xlsx_files_path, 'vehicle_sale_history.xlsx')

inventory_file = os.path.join(root_folder, "JSON", f"Inventory_{part_label}.json")

data_today_folder = os.path.join(root_folder, "Output", f"autotrader_listings_{part_label}_{todayDate}.csv")
current_date = (datetime.now() + timedelta(days=1)).strftime("%Y%m%d")

# ========================
# Logger Initialization
# ========================
logger_man = Logger(f'AutoTraders_{part_label}', root_folder=root_folder, todayDate=todayDate)
logger, log_file_name = logger_man.log, logger_man.log_file_name

logger(f"Part_Label: {part_label}", Logger.INFO)

# ========================
# Postmark Email Client
# ========================
postmark = PostmarkClient(server_token='API_KEY')


# ========================
# Load Environment Variables
# ========================

load_dotenv(dotenv_path=env_path)

def get_environment_variable(variable_name):
    """Helper function to safely fetch environment variables."""
    return os.getenv(variable_name)

# ========================
# SFTP Credentials
# ========================
SFTP_HOST_NAME = get_environment_variable('SFTP_DOWNLOAD_HOST')
SFTP_USER_NAME = get_environment_variable('SFTP_DOWNLOAD_USERNAME')
SFTP_PASSWORD = get_environment_variable('SFTP_DOWNLOAD_PASSWORD')
SFTP_PORT = get_environment_variable('SFTP_DOWNLOAD_PORT')
SFTP_FILE_UPLOAD_PATH = get_environment_variable('SFTP_DOWNLOAD_PATH')



# Global WebDriver instance
driver: WebDriver = None

# ========================
# Autotrader API URL
# ========================
api_url = ("https://www.autotrader.co.uk/at-gateway?opname=SearchResultsListingsGridQuery&opname"
           "=SearchResultsFacetsWithGroupsQuery")

# JSON payload (shortened for clarity; adjust as needed)
payload = [
    {
        "operationName": "SearchResultsListingsGridQuery",
        "variables": {
            "channel": "cars",
            "featureFlags": [],
            "page": 1,
            "sortBy": "relevance",
            "listingType": None,
        },
        "query": "",

    },
    {
        "operationName": "SearchResultsFacetsWithGroupsQuery",
        "variables": {
            "channel": "cars",
            "sortBy": "relevance",
            "featureFlags": [],
            "facets": [
                "acceleration_values",
                "aggregated_trim",
                "annual_tax_values",
                "battery_charge_time_values",
                "battery_quick_charge_time_values",
                "battery_range_values",
                "body_type",
                "boot_size_values",
                "category_tag",
                "co2_emission_values",
                "colour",
                "digital_retailing",
                "distance",
                "doors_values",
                "drivetrain",
                "engine_power",
                "engine_size",
                "seats_values",
                "seller_type",
                "style",
                "sub_style",
                "transmission",
                "ulez_compliant",
                "engine_power",
                "with_digital_retailing",
                "with_manufacturer_rrp_saving",
                "year_manufactured"
            ],
            "facetGroups": [
                "acceleration",
                "battery_range",
                "body_type",
                "boot_space",
                "category_tag",
                "charging_time",
                "co2_emissions",
                "colour",
                "keyword_search",
                "lease_price_and_terms",
                "make_and_model",
                "mileage",
                "monthly_price",
                "previously_written_off",
                "price",
                "seats",
                "seller_type",
                "tax_per_year",
                "year"
            ],
        },
        "query": ""

    }
]

class NpEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, np.integer):
            return int(obj)
        elif isinstance(obj, np.floating):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        return super(NpEncoder, self).default(obj)

def get_page(payload, page_number=0):
    """Modify the payload with a specific page number before sending API request."""
    # Deep copy payload to avoid modifying the original object
    modified_payload = json.loads(json.dumps(payload))

    if page_number:
        # Update page number inside payload (if key exists)
        for operation in modified_payload:
            if 'variables' in operation and 'page' in operation['variables']:
                operation['variables']['page'] = page_number

    return modified_payload


def initialise_driver():
    """Initialise Chrome WebDriver if not already running."""
    global driver

    # ========================
    # Chrome Driver Config
    # ========================
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--headless=new")  # if you want headless
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.7827.53 Safari/537.36")

    service = Service('/usr/local/bin/chromedriver')
    driver = webdriver.Chrome(service=service, options=chrome_options)
    logger("Chrome driver initiated...", Logger.SUCCESS)
    reopen_auto_trader()


def reopen_auto_trader():
    """Open Autotrader homepage to bypass Cloudflare protection."""
    # Load Autotrader main page
    driver.get("https://www.autotrader.co.uk")

    # Wait until the page is fully loaded (Cloudflare challenge solved)
    WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.TAG_NAME, "body")))


def get_property(property_payload, propertyName: str):
    """Fetch specific property values (e.g., fuel type, body type, etc.) from Autotrader API."""


    payloadJSON = get_page(property_payload)

    # Same XMLHttpRequest execution via JavaScript
    script = """
            var xhr = new XMLHttpRequest();
            xhr.open('POST', '%s', false);
            xhr.setRequestHeader('Content-Type', 'application/json');
            xhr.setRequestHeader('x-sauron-app-name', 'sauron-search-results-app');
            xhr.setRequestHeader('x-sauron-app-version', '93f7d286da');
            xhr.send(JSON.stringify(%s));
            if (xhr.status === 200) {
                return {
                    status: xhr.status,
                    response: xhr.responseText
                };
            } else {
                return {
                    status: xhr.status,
                    response: "Request failed with status " + xhr.status
                };
            }
            """ % (api_url, json.dumps(payloadJSON))

    # Run script & get response
    js_response = driver.execute_script(script)

    if js_response['status'] != 200:
        return []  # Return empty if failed

    try:
        # Parse response & extract requested property facet
        response_data = json.loads(js_response['response'])
        makes = [facet for facet in response_data[1]['data']['searchResults']['facets']
                 if facet['facet'] == propertyName][0]['filters'][0]['options']
        return makes
    except Exception:
        return []


def get_unique_cars(cars):
    # Deduplicate car dicts by the 'vehicle_id' field
    unique_cars = []
    seen = set()
    for car in cars:
        advertId = car.get('vehicle_id')      # Unique id for vehicle (if present)
        if advertId and advertId not in seen: # Keep only first occurrence
            seen.add(advertId)
            unique_cars.append(car)
    return unique_cars


def get_unique_advert(cars):
    # Deduplicate car dicts by the 'advertId' field (different payloads use this key)
    unique_cars = []
    seen = set()
    for car in cars:
        advertId = car.get('advertId')        # Unique id for advert (if present)
        if advertId and advertId not in seen: # Keep only first occurrence
            seen.add(advertId)
            unique_cars.append(car)
    return unique_cars


def get_cars(make: str, cars_count: int, car_payload,  retries=3, delay=5):
    # Core paginator: hits Autotrader API via browser XHR, collects listings across pages
    # initialise_driver()        # Ensure Selenium driver is ready and Cloudflare is passed
    pageNo = 0                 # Page counter (1-based in loop)
    cars = []                  # Holds all raw listings collected so far
    unique_cars = []           # Deduplicated listings (by advertId)
    iteration = 1              # Safety pass counter when counts don't match
    attempt = 0
    # while attempt < retries:
    #     try:
    while True:
        pageNo += 1
        # Create a fresh payload with the current page number
        payloadJSON = get_page(car_payload, pageNo)

        # Execute a synchronous XHR inside the browser context
        script = """
            var xhr = new XMLHttpRequest();
            xhr.open('POST', '%s', false);
            xhr.setRequestHeader('Content-Type', 'application/json');
            xhr.setRequestHeader('x-sauron-app-name', 'sauron-search-results-app');
            xhr.setRequestHeader('x-sauron-app-version', '93f7d286da');
            xhr.send(JSON.stringify(%s));
            if (xhr.status === 200) {
                return {
                    status: xhr.status,
                    response: xhr.responseText
                };
            } else {
                return {
                    status: xhr.status,
                    response: "Request failed with status " + xhr.status
                };
            };
        """ % (api_url, json.dumps(payloadJSON))

        # Run the JavaScript and capture its response
        js_response = driver.execute_script(script)

        # Stop if HTTP status != 200
        if js_response['status'] != 200:
            break

        # Parse returned JSON payload from the browser XHR
        response_data = json.loads(js_response['response'])

        # Extract relevant listing types only (exclude ads/other types)
        listings = [
            facet for facet in response_data[0]['data']['searchResults']['listings']
            if facet['type'] in ('NATURAL_LISTING', 'LEASING_LISTING')
        ]

        logger(f'Extracted {len(listings)} in page {pageNo}', Logger.SUCCESS)

        if not listings:
            # No more listings on this page:
            # 1) Deduplicate what we have
            unique_cars = get_unique_advert(cars)
            # 2) If we've reached the requested count OR retried enough times, stop
            if len(unique_cars) >= cars_count or iteration > 2:
                break
            else:
                # 3) Otherwise, restart pagination (some result sets are inconsistent)
                logger(f'Count does not match {len(unique_cars)} != {cars_count}... Continue iteration from 0..', Logger.INFO)
                pageNo = 0
                iteration += 1
                continue

        # Accumulate current page listings
        cars.extend(listings)

        # Gentle delay to look human and avoid rate limiting
        time.sleep(random.uniform(0.5, 1.2))

        # Maintain a running deduplicated view
        unique_cars = get_unique_advert(cars)

    return unique_cars


def append_to_excel(data, filename):
    # Append rows to an Excel sheet if file/sheet exist; otherwise create new file/sheet

    df = pd.DataFrame(data)
    if os.path.exists(filename):
        existing_df = pd.read_excel(filename)
        final_df = pd.concat([existing_df, df], ignore_index=True)
        final_df.to_excel(filename, index=False)
    else:
        df.to_excel(filename, index=False)


def append_to_csv(data, filename):
    """Append rows to a CSV file if it exists, otherwise create a new one."""
    df = pd.DataFrame(data)

    if os.path.exists(filename):
        # Load existing data and append new rows
        existing_df = pd.read_csv(filename)
        final_df = pd.concat([existing_df, df], ignore_index=True)
        final_df.to_csv(filename, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_ALL)
    else:
        # Create a new CSV
        df.to_csv(filename, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_ALL)

def should_resume(last_entry, **kwargs):
    # Decide whether to resume based on last saved entry matching the given state
    if not last_entry:
        return True  # No state saved => start/resume
    for key, val in kwargs.items():
        if last_entry.get(key) != val:
            return False  # Mismatch => don't resume this segment
    return True          # All keys match => safe to resume


def print_status(make_val, make_count, model_val=None, model_count=None,variant_val=None, variant_count=None, colour_val=None, colour_count=None,status="Start"):
    # Pretty console progress line with counts and timestamp
    parts = [f"Make: {make_val} ({make_count})"]
    if model_val is not None:
        parts.append(f"Model: {model_val} ({model_count})")
    if variant_val is not None:
        parts.append(f"Variant: {variant_val} ({variant_count})")
    if colour_val is not None:
        parts.append(f"Colour: {colour_val} ({colour_count})")

    logger(f"{' - '.join(parts)}, {status} at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", Logger.SUCCESS)

# ==========================
# Function to fetch cars & save immediately
# ==========================
def process_and_save(entry_vals):
    """Fetch cars for a given entry and save results with crash-safe retries."""

    # Fetch cars for the given entry values (make/model/variant/colour)
    carsData = get_cars(entry_vals['value'], entry_vals['count'], current_payload)

    # Sanity checks on the number of cars returned
    if len(carsData) > entry_vals['count']:
        logger(f"Less Count Warning: get_cars() returned {len(carsData)} items, expected {entry_vals['count']}", Logger.WARNING)
    elif len(carsData) < entry_vals['count']:
        logger(f"More Count Warning: get_cars() returned {len(carsData)} items, expected {entry_vals['count']}", Logger.WARNING)
    else:
        logger(f"Correct: get_cars() returned {len(carsData)} items, expected {entry_vals['count']}", Logger.SUCCESS)

    # ---------------- Save JSON ----------------
    category_entry = {
        "make": make['value'],
        "model": entry_vals.get('model'),
        "variant": entry_vals.get('variant'),
        "colour": entry_vals['value'] if entry_vals.get('colour_flag') else None,
        "count": len(carsData)
    }

    # Add to category_data if not already present
    if category_entry not in category_data:
        category_data.add(category_entry)

    # Save current progress to JSON immediately to prevent data loss
    with open(running_category_json_file, 'w') as f:
        json.dump(list(category_data), f, indent=2)
    logger('Append category_entry in json file.', Logger.SUCCESS)
    # --------------------SAVED AS JSON---------------------------------------------------------------------#

    # ==========================
    # Process each listing individually
    # ==========================
    skipList = []  # Used to prevent duplicate processing within this batch
    new_records = []

    for listing in carsData:
        advertId = listing.get('advertId') or ''
        listing_type = listing.get('type') or ''
        listing_make = listing.get('trackingContext', {}).get('advertContext', {}).get('make') or ''
        listing_model = listing.get('trackingContext', {}).get('advertContext', {}).get('model') or ''
        condition = listing.get('trackingContext', {}).get('advertContext', {}).get('condition') or ''
        year = listing.get('trackingContext', {}).get('advertContext', {}).get('year') or ''
        price = listing.get('trackingContext', {}).get('advertContext', {}).get('price') or listing.get('price') or ''
        sub_title = listing.get('subTitle') or ''
        attention_grabber = listing.get('attentionGrabber') or ''
        location = listing.get('location') or ''
        seller_name = listing.get('sellerName') or ''
        seller_type = listing.get('sellerType') or ''
        dealer_link_href = listing.get('dealerLink') or ''
        dealer_link = f"https://www.autotrader.co.uk{dealer_link_href}" if dealer_link_href else ''
        has_digital_retailing = listing.get('hasDigitalRetailing') or ''

        # Extract badge info
        # badges = listing.get('badges', [])
        badges = listing.get('badges') or []
        badges_mileage = next((b.get('displayText') for b in badges if b.get('type') == 'MILEAGE'), '')
        badges_registered_year = next((b.get('displayText') for b in badges if b.get('type') == 'REGISTERED_YEAR'), '')

        is_reserve_online = listing.get('hasDigitalRetailing')
        is_leasing = listing.get('type') == 'LEASING_LISTING'

        if advertId and str(advertId) not in skipList:
            record = {
                'scrape_date': scrap_date,
                'vehicle_id': str(advertId),
                'date_first_posted': datetime.strptime(advertId[:8], "%Y%m%d").strftime("%Y-%m-%d"),
                'lease_status': 1 if is_leasing else '',
                'missing_days_count': None,
                'date_last_seen': None,
                'days_to_sell': None,
                'type': str(listing_type),
                'make': str(listing_make),
                'model': str(listing_model),
                'condition': str(condition),
                'year': str(year),
                'price': str(price),
                'subTitle': str(sub_title),
                'attentionGrabber': str(attention_grabber),
                'location': str(location),
                'sellerName': str(seller_name),
                'sellerType': str(seller_type),
                'dealerLink': str(dealer_link),
                'hasDigitalRetailing': bool(has_digital_retailing),
                'badges_MILEAGE': str(badges_mileage),
                'badges_REGISTERED_YEAR': str(badges_registered_year)
            }

            new_records.append(record)

            # Track reserve online vehicles separately
            if is_reserve_online:
                reserve_online_cars_data.append(record)

    # ==========================
    # Save records to Excel
    # ==========================
    append_to_csv(new_records, csv_filename)

# ---- Function to check if vehicle is removed from AutoTrader ----
def is_vehicle_ad_removed(vehicle_id, max_retries=1, wait_seconds=0.5):
    url = f"https://www.autotrader.co.uk/product-page/v1/advert/{vehicle_id}?channel=cars&postcode=E1%207JF"
    global message

    for attempt in range(1, max_retries + 1):
        try:
            logger(f"Checking vehicle {vehicle_id} (Attempt {attempt}/{max_retries})", Logger.INFO)
            driver.get(url)
            time.sleep(wait_seconds)

            try:
                json_text = driver.find_element("tag name", "pre").text
                data = json.loads(json_text)
                message = data.get("message")
                logger(f"Message: {message}", Logger.SUCCESS)
                expected_message = f"Advert with advertId '{vehicle_id}' is not found."
                # return message.strip() == expected_message
                return message is not None and message.strip() == expected_message
                # Check alternative element text

            except Exception as inner_e:
                if attempt == max_retries:
                    try:
                        alt_message = driver.find_element(By.XPATH, "/html/body/div[3]").text.strip()
                    except:
                        alt_message = ""
                    logger(f"INNER_E: {alt_message}", Logger.WARNING)
                    # Final condition
                    return alt_message == "Advert is expired"

        except Exception as outer_e:
            logger(f"Error checking vehicle {vehicle_id} (Attempt {attempt}): {outer_e}", Logger.ERROR)
            if attempt == max_retries:
                return False

def find_latest_file(date_label):
    matching_files = [
        f for f in xlsx_files
        if os.path.basename(f).startswith(f"autotrader_listings_output_{part_label}_{date_label}.csv")
           and os.path.isfile(f)
    ]
    return max(matching_files, key=os.path.getctime) if matching_files else None

try:
    reserve_online_cars_data = []
    category_data = DictList()
    last_entry = None

    # Default: assume full resume enabled unless we detect a prior checkpoint
    resume_make = resume_model = resume_variant = resume_colour = True

    # Working buffers used during traversal/collection
    listed_cars = []   # Collected listings for current slice
    capture = False    # Toggle flag used by outer logic (set elsewhere)


    if os.path.isfile(running_category_json_file):
        # Load previous progress to resume from last checkpoint
        with open(running_category_json_file, "r") as running_category:
            data = json.load(running_category)
            category_data = DictList(data)
            last_entry = category_data[-1] if category_data else None

            # If we have a last entry, we will resume from the exact point
            if last_entry:
                resume_make = False
                resume_model = resume_variant = resume_colour = False

    # Ensure data structure is initialized if file was empty or missing
    if not category_data:
        category_data = DictList()
        last_entry = None

    # Finalize resume flags:
    # - If last_entry exists: we're resuming mid-run (disable auto-start)
    # - Else: it's a fresh run (enable all)
    if last_entry:
        resume_make = resume_model = resume_variant = resume_colour = False
    else:
        resume_make = resume_model = resume_variant = resume_colour = True  # Fresh run

    # Log current resume anchor for visibility
    logger(f"Last processed entry: {last_entry}", Logger.INFO)

    # ================================
    # VPN Helpers
    # ================================
    VPN_ROTATE = [        
        "/opt/apps/ExpressVPN/vpn-connect-ukdo.sh",
        "/opt/apps/ExpressVPN/vpn-connect-am.sh",
        "/opt/apps/ExpressVPN/vpn-connect-by.sh",
        "/opt/apps/ExpressVPN/vpn-connect-cz.sh",
        "/opt/apps/ExpressVPN/vpn-connect-uklo.sh"
    ]

    VPN_DISCONNECT_SCRIPT = "/opt/apps/ExpressVPN/vpn-disconnect-with-DNS-Reset.sh"
    vpn_index = 0


    def connect_vpn(script_path):
        try:
            logger(f"Trying VPN: {script_path}", Logger.INFO)
            subprocess.run([script_path], check=True)
            logger(f"Connected successfully with: {script_path}", Logger.SUCCESS)
            time.sleep(30)
            return True
        except subprocess.CalledProcessError as e:
            logger(f"Failed to connect using {script_path}: {e}", Logger.ERROR)
            return False

    def disconnect_vpn():
        try:
            subprocess.run([VPN_DISCONNECT_SCRIPT], check=True)
            logger("VPN disconnected.", Logger.INFO)
            time.sleep(30)
        except subprocess.CalledProcessError as e:
            logger(f"Error disconnecting VPN: {e}", Logger.ERROR)

    # # --- Initial VPN connect ---
    if not connect_vpn(VPN_ROTATE[vpn_index]):
        logger("Initial VPN connection failed, trying next...", Logger.WARNING)
        for i in range(1, len(VPN_ROTATE)):
            vpn_index = i
            if connect_vpn(VPN_ROTATE[vpn_index]):
                break
        else:
            logger("All initial VPN connection attempts failed. Exiting.", Logger.ERROR)
            raise SystemExit(1)


    max_retries = 3
    retry_delay = 100
    attempt = 0
    makes = None
    stop_processing = False
    initialise_driver()
    while attempt < max_retries and not stop_processing:
        try:
            makes = get_property(payload, 'make')

            for make in makes:
                make_val = make['value']


                # ****************************** MAIN ****************************************#
                # Start capturing makes when reaching the configured start point
                if make_val == start_make_with:
                    capture = True

                # Skip makes until we reach the starting make
                if not capture:
                    continue

                # Log which make is being processed
                logger(f"Processing: {make_val}", Logger.SUCCESS)

                # Stop processing when we reach the configured end make
                if make_val == end_make_with:
                    stop_processing = True
                    break
                # ****************************** MAIN ****************************************#

                # Optional filter skips (commented out examples for debugging/testing)
                # if make_val not in ('Zenos'):
                #     continue

                # Resume logic: unlock the make if we are resuming from last checkpoint
                if not resume_make:
                    if last_entry and make_val == last_entry.get('make'):
                        resume_make = True
                    else:
                        continue  # Skip until checkpoint is reached

                # Deep copy the payload so modifications are isolated for this make
                current_payload = copy.deepcopy(payload)

                # Add a filter for this make
                new_filter = {
                    "filter": "make",
                    "selected": make['value']
                }

                now = datetime.now()  # Capture current timestamp (used for logging)


                # ================================================
                # Logging start for the current make
                # ================================================
                logger(f"Make : {make['value']} {make['count']}, Start at : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", Logger.SUCCESS)

                # Prepare payload for make-level filtering
                car_filter_payload = copy.deepcopy(payload)
                current_payload[0]["variables"]["filters"].append(new_filter)
                current_payload[1]["variables"]["filters"].append(new_filter)

                # ==========================================================================
                # Deep filtering: model -> variant -> colour if counts > 2000
                # ==========================================================================
                if make['count'] > 2000:
                    car_filter_payload[0]["variables"]["filters"].append(new_filter)
                    car_filter_payload[1]["variables"]["filters"].append(new_filter)

                    # Fetch all models for this make
                    models = get_property(car_filter_payload, 'model')

                    for model in models:
                        model_val = model['value']

                        # Resume logic for model level
                        if not resume_model:
                            if should_resume(last_entry, make=make_val, model=model_val):
                                resume_model = True
                            else:
                                continue

                        model_filter = {"filter": "model", "selected": model_val}
                        print_status(make_val, make['count'], model_val, model['count'], status="Start")

                        current_payload[0]["variables"]["filters"].append(model_filter)
                        current_payload[1]["variables"]["filters"].append(model_filter)

                        if model['count'] > 2000:
                            car_filter_payload[0]["variables"]["filters"].append(model_filter)
                            car_filter_payload[1]["variables"]["filters"].append(model_filter)

                            # Fetch variants for this model
                            variants = get_property(car_filter_payload, 'aggregated_trim')

                            for variant in variants:
                                variant_val = variant['value']

                                # Resume logic for variant level
                                if not resume_variant:
                                    if should_resume(last_entry, make=make_val, model=model_val, variant=variant_val):
                                        resume_variant = True
                                    else:
                                        continue

                                variant_filter = {"filter": "aggregated_trim", "selected": variant_val}
                                print_status(make_val, make['count'], model_val, model['count'],
                                             variant_val, variant['count'], status="Start")

                                current_payload[0]["variables"]["filters"].append(variant_filter)
                                current_payload[1]["variables"]["filters"].append(variant_filter)

                                if variant['count'] > 2000:
                                    car_filter_payload[0]["variables"]["filters"].append(variant_filter)
                                    car_filter_payload[1]["variables"]["filters"].append(variant_filter)

                                    # Fetch colours for this variant
                                    colours = get_property(car_filter_payload, 'colour')
                                    for colour in colours:
                                        colour_val = colour['value']

                                        # Resume logic for colour level
                                        if not resume_colour:
                                            if should_resume(last_entry, make=make_val, model=model_val, variant=variant_val,
                                                             colour=colour_val):
                                                resume_colour = True
                                            else:
                                                continue

                                        colour_filter = {"filter": "colour", "selected": colour_val}
                                        print_status(make_val, make['count'], model_val, model['count'],
                                                     variant_val, variant['count'], colour_val, colour['count'], status="Start")

                                        # Apply colour filter and process cars
                                        current_payload[0]["variables"]["filters"].append(colour_filter)
                                        current_payload[1]["variables"]["filters"].append(colour_filter)

                                        process_and_save({
                                            "value": colour_val,
                                            "count": colour['count'],
                                            "model": model_val,
                                            "variant": variant_val,
                                            "colour_flag": True
                                        })

                                        # Remove colour filter after processing
                                        current_payload[0]["variables"]["filters"].pop()
                                        current_payload[1]["variables"]["filters"].pop()

                                        print_status(make_val, make['count'], model_val, model['count'],
                                                     variant_val, variant['count'], colour_val, colour['count'], status="End")

                                    # Pop variant-level filters after all colours
                                    car_filter_payload[0]["variables"]["filters"].pop()
                                    car_filter_payload[1]["variables"]["filters"].pop()

                                    # Reset colour resume flag
                                    resume_colour = True

                                else:
                                    # If variant count <= 2000, process directly without colour
                                    process_and_save({
                                        "value": variant_val,
                                        "count": variant['count'],
                                        "model": model_val,
                                        "variant": variant_val
                                    })

                                    print_status(make_val, make['count'], model_val, model['count'],
                                                 variant_val, variant['count'], status="End")

                                # Pop variant filter after processing
                                current_payload[0]["variables"]["filters"].pop()
                                current_payload[1]["variables"]["filters"].pop()

                            # Pop model filters after variants
                            car_filter_payload[0]["variables"]["filters"].pop()
                            car_filter_payload[1]["variables"]["filters"].pop()

                            # Reset variant & colour resume flags
                            resume_variant = True
                            resume_colour = True

                        else:
                            # Process model directly if count <= 2000
                            process_and_save({
                                "value": model_val,
                                "count": model['count'],
                                "model": model_val
                            })

                            print_status(make_val, make['count'], model_val, model['count'], status="End")

                        # Pop model filters after processing
                        current_payload[0]["variables"]["filters"].pop()
                        current_payload[1]["variables"]["filters"].pop()

                    # Pop make filters after all models
                    car_filter_payload[0]["variables"]["filters"].pop()
                    car_filter_payload[1]["variables"]["filters"].pop()

                    # Reset model-related resume flags
                    resume_model = resume_variant = resume_colour = True

                else:
                    # If make count <= 2000, process make directly without model/variant/colour filtering
                    process_and_save({
                        "value": make_val,
                        "count": make['count']
                    })

                    print_status(make_val, make['count'], status="End")

                # Pop make filters after processing
                current_payload[0]["variables"]["filters"].pop()
                current_payload[1]["variables"]["filters"].pop()

                # Reset all resume flags for next make
                resume_make = resume_model = resume_variant = resume_colour = True

                logger(f"Completed make '{make_val}' and ready for next make\n", Logger.SUCCESS)

        except Exception as e:
            attempt += 1
            logger(f"Attempt {attempt} failed while fetching makes: {str(e)}", Logger.WARNING)

            if attempt < max_retries:
                # Reinitialize driver
                try:
                    driver.quit()
                except Exception:
                    pass

                logger(f"Retrying in {retry_delay} seconds...", Logger.INFO)
                time.sleep(retry_delay)



                disconnect_vpn()

                vpn_index = (vpn_index + 1) % len(VPN_ROTATE)
                next_vpn = VPN_ROTATE[vpn_index]

                # 3) Connect new VPN
                if connect_vpn(next_vpn):
                    logger(f"Switched to VPN: {next_vpn}", Logger.SUCCESS)
                else:
                    logger(f"VPN connection failed for {next_vpn}, will retry with next option.", Logger.ERROR)


                # Reload last checkpoint to continue safely
                if os.path.isfile(running_category_json_file):
                    with open(running_category_json_file, "r") as running_category:
                        data = json.load(running_category)
                        category_data = DictList(data)
                        last_entry = category_data[-1] if category_data else None

                        # Resume flags
                        if last_entry:
                            resume_make = resume_model = resume_variant = resume_colour = False
                        else:
                            resume_make = resume_model = resume_variant = resume_colour = True

                initialise_driver()
            else:
                # Exceeded retries
                logger("Max retries reached while fetching makes. Aborting run.", Logger.ERROR)
                raise

    # ================================================
    # Read Excel file containing today's listings
    # ================================================
    df = pd.read_csv(csv_filename)
    all_cars_data = df.to_dict(orient='records')

    # Remove duplicate cars based on vehicle_id
    uniqueCars = get_unique_cars(all_cars_data)
    logger(f"Total no of UniqueCars: {len(uniqueCars)}", Logger.INFO)


    # Use DictList utility for easy dictionary management and indexing
    current_listings_list = DictList(list_of_dicts=uniqueCars, index_column_name='vehicle_id')
    current_listings_list.save_as_file(output_filename)

    # Reload to replace NaNs with None for consistency
    df = pd.read_csv(output_filename)
    df = df.replace({np.nan: None})
    all_cars_data = df.to_dict(orient='records')
    no_of_vehicle_in_platform = len(all_cars_data)

    # ========================================================================
    # Compare with previous day's listings to detect sold vehicles
    # ========================================================================
    try:
        # if all_cars_data:
        current_listings_list = DictList(list_of_dicts=all_cars_data, index_column_name='vehicle_id')
        previous_listings_list = DictList(index_column_name='vehicle_id')

        # Try yesterday first, then day-before-yesterday
        latest_xlsx = find_latest_file(yesterday_output_file)

        if not latest_xlsx:
            print(f"No file found for {yesterday_output_file}, checking {day_before_yesterday_output_file}")
            latest_xlsx = find_latest_file(day_before_yesterday_output_file)

        # Log results
        if latest_xlsx:
            print(f"Valid last modified excel file: {latest_xlsx}")
        else:
            print("No valid xlsx files found for the last two days.")

        print(f"Latest xlsx file found: {latest_xlsx}")

        # Load previous listings
        if latest_xlsx:
            previous_listings_list.read_dict_list_from_csv(latest_xlsx)

            # Filter only unsold vehicles from previous data
            df_previous_full = pd.DataFrame(previous_listings_list)
            df_previous = df_previous_full[
                (df_previous_full.get('date_last_seen').isna()) &
                (df_previous_full.get('days_to_sell').isna())
                ].copy()
        else:
            df_previous = pd.DataFrame(previous_listings_list)

        # Read your Excel file
        df = pd.read_csv(csv_filename)

        # Ensure the column exists
        if 'hasDigitalRetailing' in df.columns:
            # Convert column to consistent boolean values
            df['hasDigitalRetailing'] = df['hasDigitalRetailing'].astype(str).str.lower().isin(['true', '1'])

            # Count True values
            vehicle_reserve_online = df['hasDigitalRetailing'].sum()
        else:
            vehicle_reserve_online = 0


        # Initialize counters
        sold_Vehicles = 0
        vehicle_days_to_sell = 0
        vehicles_on_platform = len(current_listings_list)
        vehicle_reserve_old = len(reserve_online_cars_data)
        logger(f"vehicle_reserve_old: {vehicle_reserve_old}", Logger.INFO)
        logger(f"vehicle_reserve_online: {vehicle_reserve_online}", Logger.INFO)

        # ========================================================================
        # Detect sold vehicles by comparing previous and current listings
        # ========================================================================

        df_current = pd.DataFrame(current_listings_list)
        df_current['vehicle_id'] = df_current['vehicle_id'].astype(str)
        df_previous['vehicle_id'] = df_previous['vehicle_id'].astype(str)
        df_current.sort_values(by='vehicle_id', inplace=True)
        df_previous.sort_values(by='vehicle_id', inplace=True)

        sold_ids = set(df_previous['vehicle_id']) - set(df_current['vehicle_id'])
        df_sold = df_previous[df_previous['vehicle_id'].isin(sold_ids)].copy()
        logger(f"Found {len(df_sold)} sold vehicles.", Logger.SUCCESS)

        # Update sold vehicles with last seen date and days to sell
        if not df_sold.empty:
            df_sold['date_last_seen'] = yesterdayDateInDash
            df_sold['days_to_sell'] = (
                    pd.to_datetime(yesterdayDateInDash) - pd.to_datetime(df_sold['date_first_posted'])
            ).dt.days
            df_sold['missing_days_count'] = None

            df_final = pd.concat([df_current, df_sold], ignore_index=True)
        else:
            df_final = df_current.copy()

        # Save final Excel with current + sold vehicles
        df_final.to_csv(csv_filename, index=False)
        logger(f"segregate sold list excel saved with {len(df_final)} total listings: {csv_filename}", Logger.SUCCESS)


        # ========================================================================
        # Function to check if a vehicle has been removed from AutoTrader
        # ========================================================================
        df = pd.read_csv(csv_filename, low_memory=False)

        save_every = 100

        # ---- Check each sold vehicle and set missing_days_count if applicable ----
        for idx, row in df.iterrows():
            if pd.notna(row.get('date_last_seen')) and pd.notna(row.get('days_to_sell')):
                vehicle_id = str(int(float(row['vehicle_id'])))
                removed = is_vehicle_ad_removed(vehicle_id)

                current_missing = row.get('missing_days_count')
                current_missing_count = int(current_missing) if pd.notna(current_missing) and str(
                    current_missing).isdigit() else 0

                if removed:
                    df.at[idx, 'missing_days_count'] = ''
                    logger(f"Checked {vehicle_id}: REMOVED from AutoTrader", Logger.SUCCESS)
                else:
                    df.at[idx, 'missing_days_count'] = current_missing_count + 1
                    logger(f"Checked {vehicle_id}: Still live -> missing_days_count = {current_missing_count + 1}", Logger.SUCCESS)

                if idx % save_every == 0:
                    df.to_csv(csv_filename, index=False)
                    logger(f"Progress saved at row {idx}", Logger.SUCCESS)

        # ---- Save updated Excel with validated missing_days_count ----
        df.to_csv(csv_filename, index=False)
        logger(f"Missing days count check excel saved: {csv_filename}", Logger.SUCCESS)

        # ========================
        # Update vehicle sale history and inventory
        # ========================

        current_inventory = DictList(index_column_name='scrape_date')

        # Load existing inventory if exists
        if os.path.isfile(xlsx_history_files):
            current_inventory.read_dict_list_from_xlsx(xlsx_history_files)

        df = pd.read_csv(csv_filename)

        # Calculate average days to sell
        average_days = df['days_to_sell'].mean()
        avg_lead_time_for_vehicles_sold = round(average_days)

        # ========================
        # Also save inventory as JSON for quick reference
        # ========================

        new_entry = {
            "scrape_date": scrap_date,
            "no_of_vehicles_on_platform": no_of_vehicle_in_platform,
            "no_vehicles_reserve_online": vehicle_reserve_online,
            "avg_lead_time_for_vehicles_sold": avg_lead_time_for_vehicles_sold
        }

        # Load existing JSON if exists
        if os.path.exists(inventory_file):
            with open(inventory_file, "r") as f:
                inventory_data = json.load(f)
        else:
            inventory_data = []

        inventory_data.append(new_entry)

        # Save updated JSON inventory
        with open(inventory_file, "w") as f:
            json.dump(inventory_data, f, indent=4, cls=NpEncoder)
        logger("Inventory updated as JSON file.", Logger.SUCCESS)

    except KeyError:
        logger("No listing data was collected", Logger.WARNING)

    disconnect_vpn()
    time.sleep(10)
    
    
    # ================================================
    # Upload Excel and JSON files to SFTP
    # ================================================

    # Generate remote folder for today's date
    remote_dir = f"{SFTP_FILE_UPLOAD_PATH}{current_date}"

    # Initialize SFTP connection
    transport = paramiko.Transport((SFTP_HOST_NAME, int(SFTP_PORT)))
    transport.connect(username=SFTP_USER_NAME, password=SFTP_PASSWORD)
    sftp = paramiko.SFTPClient.from_transport(transport)

    # Create remote folder if it doesn't exist
    try:
        sftp.listdir(remote_dir)
    except IOError:
        sftp.mkdir(remote_dir)

    # Upload Excel file
    if os.path.exists(data_today_folder):
        excel_filename = os.path.basename(data_today_folder)
        excel_remote_path = f"{remote_dir}/{excel_filename}"
        sftp.put(data_today_folder, excel_remote_path)
        logger(f"Excel uploaded: {excel_remote_path}", Logger.SUCCESS)
    else:
        logger(f"Excel file not found, skipping upload: {data_today_folder}", Logger.WARNING)

    # Upload JSON inventory file
    if os.path.exists(inventory_file):
        json_filename = os.path.basename(inventory_file)
        json_remote_path = f"{remote_dir}/{json_filename}"
        sftp.put(inventory_file, json_remote_path)
        logger(f"JSON uploaded: {json_remote_path}", Logger.SUCCESS)
    else:
        logger(f"JSON file not found, skipping upload: {inventory_file}", Logger.WARNING)

    # Close SFTP connection
    sftp.close()
    transport.close()

    postmark.emails.send(
        From='XYX@gmail.com',
        To='XXX@gmail.com',
        Subject=f"Man's group AutoTraders_{part_label} | SUCCESS [Server_name]",
        HtmlBody=f'AutoTraders_{part_label} output file generated successfully.',
        Attachments=[log_file_name, inventory_file]
    )
except Exception as e:
    error_details = traceback.format_exc()
    logger(f"Unexpected error: {error_details}", Logger.ERROR)
    postmark.emails.send(
        From='XYX@gmail.com',
        To='XXX@gmail.com',
        Subject=f"Man's group AutoTraders_{part_label} | Alert [Server_name]",
        HtmlBody=f'AutoTraders_{part_label} failed. Check log file to view the issue.',
        Attachments=[log_file_name]
    )

finally:
    driver.quit()

    # --- Always disconnect VPN at the end ---
    disconnect_vpn()