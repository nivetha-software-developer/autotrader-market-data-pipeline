# 🚘 AutoTrader UK Market Data ETL Pipeline

[![Python](https://img.shields.io/badge/Python-3.9%2B-blue?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![Selenium](https://img.shields.io/badge/Selenium-4.0%2B-43B02A?style=flat-square&logo=selenium&logoColor=white)](https://www.selenium.dev/)
[![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-150458?style=flat-square&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![Status](https://img.shields.io/badge/Pipeline-Production--Ready-success?style=flat-square)](#)

An enterprise-grade, resilient data extraction and ETL pipeline built with Python and Selenium to scrape, clean, deduplicate, and analyze car listings from AutoTrader UK.

---

## 🏗️ Architecture & Core Features

```mermaid
flowchart TD
    A[Start Pipeline] --> B[Initialize Headless Chrome WebDriver]
    B --> C[ExpressVPN Auto-Rotation & IP Switch]    
    C --> D[Bypass Cloudflare via Session Warmup]
    D --> E[Inject Synchronous XHR for Gateway Query]
    E --> F{Facets > 2000 Listings?}
    F -- Yes --> G[Recursive Filter Breakdown: Make -> Model -> Variant -> Colour]
    F -- No --> H[Extract Listing Details & Pagination]
    G --> H
    H --> I[Checkpoint State to JSON - Fault Recovery]
    I --> J[Deduplicate Listings & Compute Market Metrics]
    J --> K[Detect Sold Vehicles & Days-To-Sell Statistics]
    K --> L[Upload Output & Inventory Reports to SFTP]
    L --> M[Postmark Email Notification with Attached Logs]

```
---

## Key Engineering Capabilities
- 🛡️ Anti-Bot & Cloudflare Bypass: Executes headless browser sessions and injects synchronous JavaScript XMLHttpRequest (XHR) directly into the browser execution context.  
- 🔄 Automated VPN Rotation: Integrates shell-based ExpressVPN script execution to rotate IPs automatically on connection failures or rate limits.  
- 🌲 Dynamic Facet Partitioning: Handles AutoTrader API page limits by dynamically splitting query payloads down to Make -> Model -> Variant -> Colour when result thresholds exceed 2,000 items.  
- 💾 Crash-Safe Checkpointing: Progress is continuously written to state JSON files, allowing execution recovery from exact interruption nodes.  
- 📊 Market Intelligence Analytics: Tracks vehicle listing lifespans, missing days counts, sold statuses, and platform inventory metrics over time.  
- 📤 Enterprise Delivery & Alerts: Automatically uploads datasets to remote SFTP servers and dispatches operational logs via Postmark API.

---

## 🛠️ Tech Stack
- Core Language: Python 3.9+  
- Automation & Scraping: Selenium WebDriver, ChromeDriverManager  
- Data Manipulation & ETL: Pandas, NumPy, JSON, Openpyxl, CSV/Excel  
- Network & Infrastructure: Paramiko (SFTP), Postmarker API, ExpressVPN CLI Integration

---

## ⚙️ Configuration & Setup
#### 1. Prerequisites
   - Python 3.9 or higher installed
   - Google Chrome browser & ChromeDriver matching your version
   - ExpressVPN Linux CLI (if automated VPN rotation is required)
     
#### 2. Installation
Clone the repository and install required packages:

```Bash
git clone [https://github.com/nivethamanoharan/autotrader-market-data-pipeline.git](https://github.com/nivethamanoharan/autotrader-market-data-pipeline.git)
cd autotrader-market-data-pipeline
pip install -r requirements.txt
```

#### 3. Environment Variables
Create a .env file in the root directory:

```Code snippet
SFTP_DOWNLOAD_HOST=sftp.example.com
SFTP_DOWNLOAD_USERNAME=your_username
SFTP_DOWNLOAD_PASSWORD=your_password
SFTP_DOWNLOAD_PORT=22
SFTP_DOWNLOAD_PATH=/remote/market_data/
```

---

## 🚀 Running the Pipeline
To execute the market extraction and data synchronization pipeline:

```Bash
python VPN_AUTOTRADER.py
```

---

## 📊 Sample Output Schema

| Field Name | Type | Descriptions | 
|------------|------|--------------|
| Scrape_date | String |Date of pipeline execution (YYYY-MM-DD)  
| vehicle_id | String | Unique AutoTrader vehicle listing identifier |  
| date_first_posted | String | Derived listing publish date |  
| make / model | String | Vehicle manufacturer and model designation |  
| price | String | Advertised vehicle price |  
| days_to_sell | Integer | Estimated market lead time upon listing removal |
| hasDigitalRetailing | Boolean | Reserve-online status indicator |

---

## 📧 Author & Connect
#### Nivetha Manoharan
> Software Developer (Data Engineering & Automation)
- 💼 LinkedIn: linkedin.com/in/nivethamanoharan  
- ✉️ Email: nivemanoharan2001@gmail.com  
- 📍 Status: Open to relocation to UAE | Immediate Availability  
