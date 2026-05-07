# E-Commerce Automated Price Scraper 🛒🤖

A robust, fully automated Python script designed to scrape product prices from major Egyptian e-commerce platforms (Dubai Phone, Dream 2000, 2B). It bypasses modern anti-bot protections (like Cloudflare), processes the data using Pandas, and automatically emails a daily Excel report.

## 🚀 Features
* **Anti-Bot Bypass:** Utilizes `undetected-chromedriver` to seamlessly bypass Cloudflare and other web application firewalls.
* **Data Processing:** Uses `pandas` and `openpyxl` to read input URLs, append newly scraped prices, and generate clean, formatted Excel reports.
* **Automated Delivery:** Integrates with Gmail's SMTP server to send the final daily report directly to the client's inbox.
* **Task Scheduling:** Runs automatically at a predefined time (e.g., 08:00 AM) using the `schedule` library.

## 🛠️ Tech Stack
* Python 3.x
* Selenium / Undetected-Chromedriver
* Pandas
* SMTPlib

## 📝 How to Use
1. **Clone the repository:**
   ```bash
   git clone [https://github.com/yourusername/egyptian-ecommerce-price-scraper.git](https://github.com/yourusername/egyptian-ecommerce-price-scraper.git)
