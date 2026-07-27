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
