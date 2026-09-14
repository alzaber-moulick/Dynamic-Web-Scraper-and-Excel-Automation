<img width="1200" height="896" alt="automated web scrapping thumbnails" src="https://github.com/user-attachments/assets/43fe7b16-d39c-463c-bbff-a9ad310a37e9" />
# Dynamic-Web-Scraper-and-Excel-Automation
A dynamic Python web scraper using Selenium, BeautifulSoup, and Pandas to extract multi-level data into multi-sheet Excel reports.
# 🏏 Dynamic Web Scraper & Multi-Tab Excel Data Automation

## 📌 Business Overview & Solution
Modern web applications heavily rely on dynamic JavaScript rendering, making traditional web scraping methods ineffective. This project provides an automated, end-to-end Python scraping pipeline built for the **Crickslab Match Central** platform.

It solves the problem of extracting nested, multi-level player statistics across dynamically loaded DOM elements. The automation script navigates through dynamic directory listings, parses individual player profiles, extracts **32 granular career parameters**, and automatically organizes the output into a formatted multi-sheet Excel report (`.xlsx`).

---

## 📸 Project Showcase & Visual Output

<p align="center">
  <img src="automated%20web%20scrapping%20thumbnails.jpeg" alt="Automated Web Scraping Dashboard Showcase" width="100%">
</p>

---

## 🛠️ Key Capabilities & Technical Features

- **Dynamic JavaScript Handling:** Deploys Selenium WebDriver in headless mode to navigate client-side JavaScript execution, dynamic page scrolling, and asynchronous data rendering.
- **2-Step Scraping Workflow:**
  1. **Link Harvesting:** Automatically scrolls through directory listings to aggregate all unique player profile URLs.
  2. **Granular Extraction:** Visits each profile sequentially to extract demographic, batting, bowling, and fielding metrics using BeautifulSoup & Regex.
- **Fault-Tolerant Data Parsing:** Built with Regular Expressions (`re`) to handle missing data fields, string sanitation, dynamic table structures, and type casting safely.
- **Automated Multi-Sheet Excel Export:** Uses `pandas` and `openpyxl` to automatically segment and export extracted raw data into categorized worksheets for business intelligence and client review.

---

## 📊 Extracted Data Schema (32 Fields)

The scraper extracts, cleans, and structures 32 distinct data points grouped into:
1. **Profile Demographics:** Player Name, Role, Nationality, City, Batting Style, Bowling Style, Profile URL.
2. **Career Summaries:** Total Matches, Total Runs, Total Wickets.
3. **Batting Metrics:** Matches, Innings, Runs, High Score, Batting Average, Strike Rate, 50s, 100s, 4s, 6s.
4. **Bowling & Fielding Metrics:** Bowling Matches, Innings, Overs, Maidens, Runs Conceded, Wickets, Economy, Bowling Average, Bowling SR, 5W, Catches, Run Outs.

---

## 📂 Multi-Tab Excel Workbook Structure

The generated Excel file (`Crickslab_Players_Data.xlsx`) contains 4 structured tabs:
- 📄 **`All Players`:** Complete master database containing all 32 metrics for all parsed profiles.
- 🏏 **`Batsmen`:** Auto-filtered dataset focusing on specialist batters.
- 🎯 **`Bowlers`:** Auto-filtered dataset highlighting specialist bowlers.
- ⚡ **`All Rounders`:** Filtered dataset for players with dual career statistics.

---

## 🧰 Tech Stack & Libraries

- **Automation:** `Selenium WebDriver`, `webdriver-manager`
- **Parsing & Extraction:** `BeautifulSoup4`, `re` (Regular Expressions)
- **Data Engineering & Output:** `Pandas`, `OpenPyXL`
- **Language:** `Python 3.x`

---

## 🚀 How to Run the Script

### 1. Clone the Repository
```bash
git clone [https://github.com/your-username/Dynamic-Web-Scraper-and-Excel-Automation.git](https://github.com/your-username/Dynamic-Web-Scraper-and-Excel-Automation.git)
cd Dynamic-Web-Scraper-and-Excel-Automation
