import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import schedule
import time
import os

# Import openpyxl styles for formatting the Excel file
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ----------------- 1. Browser Setup -----------------
def setup_driver():
    """
    Initializes undetected-chromedriver.
    Note: version_main should match the host's Chrome version.
    """
    options = uc.ChromeOptions()
    options.headless = False  # Headless=False is safer against Cloudflare detection

    # Update version_main if the browser version differs
    driver = uc.Chrome(options=options, version_main=147)
    driver.maximize_window()
    return driver


# ----------------- 2. Scraping Functions -----------------
def get_price(driver, url, css_selector):
    """Fetches price from a given URL using a CSS selector."""
    if pd.isna(url) or str(url).strip() == "":
        return "N/A"

    try:
        driver.get(url)
        # Wait for the element to be visible
        price_elem = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, css_selector))
        )
        return price_elem.text.strip()
    except Exception:
        return "Error / Not Found"


# ----------------- 3. Email Function -----------------
def send_email_with_report(file_path):
    """Attaches the Excel report and sends it via Gmail SMTP."""
    # TODO: Replace with your actual credentials before local execution
    sender_email = "your_email@gmail.com"
    app_password = "your_app_password_here"
    receiver_email = "client_email@example.com"

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg["Subject"] = f"Automated Daily Price Report - {datetime.date.today()}"

    body = (
        "Hello,\n\n"
        "Please find attached the daily product price report extracted from the targeted e-commerce platforms.\n\n"
        "Best regards,\n"
        "Python Automation Script"
    )
    msg.attach(MIMEText(body, "plain", "utf-8"))

    try:
        with open(file_path, "rb") as f:
            attach = MIMEApplication(f.read(), _subtype="xlsx")
            attach.add_header(
                "Content-Disposition",
                "attachment",
                filename=os.path.basename(file_path),
            )
            msg.attach(attach)

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, app_password)
        server.send_message(msg)
        server.quit()
        print(f"[{datetime.datetime.now()}] Email sent successfully.")
    except Exception as e:
        print(f"[{datetime.datetime.now()}] Failed to send email: {e}")


# ----------------- 4. Main Automation Logic -----------------
def run_daily_scraper():
    print(f"\n[{datetime.datetime.now()}] Starting daily price extraction...")

    input_file = "warehouseStockspricesfinal.xlsx"
    try:
        # Reading from the source Excel file
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"Error reading file {input_file}: {e}")
        return

    driver = setup_driver()

    dubaiphone_prices = []
    dream2000_prices = []
    twob_prices = []

    for index, row in df.iterrows():
        product_name = row.get("Product_Name", f"Product {index+1}")
        print(f"[{index+1}/{len(df)}] Checking: {product_name}...")

        # Selectors for target websites (Ensure Excel columns match these keys exactly)
        dp_price = get_price(
            driver, row.get("DubaiPhone_URL", ""), ".price-wrapper .price"
        )
        dr_price = get_price(
            driver, row.get("Dream2000_URL", ""), ".product-info-price .price"
        )
        b2_price = get_price(driver, row.get("2B_URL", ""), ".price-box .price")

        dubaiphone_prices.append(dp_price)
        dream2000_prices.append(dr_price)
        twob_prices.append(b2_price)

        time.sleep(2)  # Anti-throttle delay

    driver.quit()

    # Append results to the dataframe
    df["Dubai Phone (Update)"] = dubaiphone_prices
    df["Dream 2000 (Update)"] = dream2000_prices
    df["2B (Update)"] = twob_prices

    # File naming with current date
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    output_filename = f"Daily_Prices_{today_str}.xlsx"

    try:
        # Use pandas ExcelWriter to format the output using openpyxl
        with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Prices")
            workbook = writer.book
            worksheet = writer.sheets["Prices"]

            # Define styling properties
            header_font = Font(bold=True, color="FFFFFF")
            # Blue background for headers
            header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            center_alignment = Alignment(horizontal="center", vertical="center")
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                 top=Side(style='thin'), bottom=Side(style='thin'))

            # Apply styling to header row
            for col_num, value in enumerate(df.columns.values):
                cell = worksheet.cell(row=1, column=col_num + 1)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center_alignment
                cell.border = thin_border

            # Auto-adjust column width and style all data cells
            for col in worksheet.columns:
                max_length = 0
                column_letter = col[0].column_letter # Get the column letter (e.g., 'A', 'B')
                
                for cell in col:
                    # Apply borders and alignment to every cell
                    if cell.row != 1: # Header is already styled
                        cell.border = thin_border
                        cell.alignment = center_alignment
                    
                    # Calculate max length for auto-fitting
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                # Add a little padding to the width
                adjusted_width = (max_length + 4)
                worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"Report saved and formatted successfully: {output_filename}")
        send_email_with_report(output_filename)
        
    except Exception as e:
        print(f"Error saving Excel file: {e}")


# ----------------- 5. Scheduler -----------------
if __name__ == "__main__":
    print("Scheduler active. Waiting for 08:00 AM...")

    # Job scheduling
    schedule.every().day.at("08:00").do(run_daily_scraper)

    # Uncomment the line below to run immediately for testing purposes
    # run_daily_scraper()

    # Continuous loop to check the schedule
    while True:
        schedule.run_pending()
        time.sleep(30)
