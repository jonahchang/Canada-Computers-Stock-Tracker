# Canada Computers Stock Tracker (CCST)

A Python-based automation tool that tracks **real-time inventory** of products across all Canada Computers stores, including **online availability**, and compiles weekly stock reports into **Excel spreadsheets** for historical trend analysis.  

This project combines **web automation, data engineering, and visualization** into a complete pipeline for market research, business intelligence, and supply chain monitoring.  

---

## Features  

- **Automated Stock Collection**  
  - Uses Selenium WebDriver to scrape product pages directly from Canada Computers.  
  - Detects **per-store inventory counts** (`0`, exact number, or `10+`) by parsing DOM class markers (`bg-0000001` for 0, `bg-E3E9F8` for in-stock).  
  - Tracks **online shipping availability** (Available to Ship / Sold Out Online).  
  - Supports arbitrary ordering of stores in the webpage while outputting results in a consistent predefined order.  

- **Excel Integration**  
  - Generates weekly stock reports in a single `.xlsx` workbook.  
  - Automatically formats category and product tables with:  
    - **Merged category cells** spanning product rows  
    - Per-store stock columns  
    - Online availability  
    - **Individual product totals**  
    - **Category totals**  
    - **Out-of-stock counts**  
  - Color-coded highlighting of weekly changes:  
    - 🟩 **Green** → Stock increase  
    - 🟨 **Yellow** → Stock decrease  
    - 🟦 **Blue** → Availability appeared/disappeared  

- **Charts & Visualization**  
  - Multi-line charts showing **model stock trends** across weeks.  
  - Category-level summaries for broader inventory insight.  
  - Store-level summaries to analyze **regional distribution**.  

- **Product Management UI**  
  - Tkinter-based editor for adding/removing tracked products.  
  - JSON-based product lists for persistence and version control.  
  - Reset option to restore original datasets.  

- **Debugging & Transparency**  
  - Saves raw scraping logs (store element HTML vs extracted stock values).  
  - Console debug output links each **store → matched HTML → parsed stock**.  
  - Debug history appended to file for reference across runs.  

---

## Tech Stack  

- **Python 3.11+**  
- **Selenium + WebDriver Manager** → Browser automation & scraping  
- **OpenPyXL** → Excel reporting, formatting, and charting  
- **Tkinter** → GUI for product management  
- **JSON** → Product data storage  

---

## Example Output  

- Weekly Excel sheet (`Canada Computers Stock Tracker.xlsx`) with:  
  - **Category totals** (Chassis, Coolers, Externals, Power Supplies, etc.)  
  - **Product breakdowns** with online status + store stock levels  
  - **Store totals** across all tracked products  

### Excel Table Example  

| Product Category | Model              | Online | Surrey | Richmond | Burnaby | … | INDIVIDUAL TOTALS | CATEGORY TOTALS | OUT OF STOCK |
|------------------|--------------------|--------|--------|----------|---------|---|-------------------|-----------------|---------------|
| Chassis          | GT301/BLK/ARGB FAN | Yes    | 3      | 1        | 0       | … | 12                | 56              | 5             |
|                  | GT502/WHT/TG       | No     | 0      | 0        | 2       | … | 4                 |                 | 8             |

---

## Business Value  

This tool automates the otherwise **manual, time-intensive process** of tracking inventory across dozens of Canada Computers locations.  

Potential applications:  
- Competitive benchmarking  
- Regional demand forecasting  
- Retail supply chain monitoring  
- Consumer trend analysis  

---

## How It Works  

1. **Start the script**  
   - Choose whether to create a new Excel workbook or update an existing one.  
   - Enter the **week number** (e.g. `9` → `WK9` sheet).  

2. **Stock Tracker runs**  
   - Opens each product page.  
   - Detects store names + stock counts, plus online availability.  
   - Writes results into Excel with formatting and merged category headers.  

3. **Data Analysis**  
   - Highlights changes week-to-week.  
   - Produces model, category, and store-level summaries.  

4. **Output**  
   - Updated Excel workbook with tables, charts, and debug logs.  

---

## Project Structure  

- **CCST.py** → Main stock tracker script  
- **CC_products.json** → Editable product list  
- **CC_original_products.json** → Original baseline product list  
- **Canada Computers Stock Tracker.xlsx** → Excel workbook with weekly sheets, trends, and charts  

---

## Quick Start  

```bash
# Clone repo
git clone https://github.com/jonahchang/Canada-Computers-Stock-Tracker.git
cd Canada-Computers_Stock_Tracker

# Install dependencies
pip install -r requirements.txt

# Run tracker
python CCST.py
