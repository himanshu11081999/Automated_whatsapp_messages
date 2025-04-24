from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager  # ✅ Required for ChromeDriverManager
import pandas as pd
import time
import urllib.parse
import os

# Load Excel File
file_path = r"E:\CMC\Demo.xlsx"
df = pd.read_excel(file_path, engine='openpyxl')

# Validate required columns
if "Phone" not in df.columns or "Message" not in df.columns or "Name" not in df.columns:
    raise ValueError("Excel must contain 'Phone', 'Name', and 'Message' columns!")

# Set path to chromedriver.exe (make sure version matches your Chrome)
chromedriver_path = r"F:\Softwares\Java\cd\chromedriver-win64\chromedriver-win64\chromedriver.exe"
if not os.path.exists(chromedriver_path):
    raise FileNotFoundError("Chromedriver path is invalid. Check the file location.")

# Initialize Chrome with the updated Service class
options = Options()
options.add_argument("--start-maximized")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# Open WhatsApp Web
driver.get("https://web.whatsapp.com")
print("Please scan the QR code in the opened browser window...")

# Wait until WhatsApp Web is ready
WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.ID, "side"))
)

# Iterate through rows in Excel
for index, row in df.iterrows():
    phone_number = str(row["Phone"]).strip()
    name = str(row["Name"]).strip()

    # Fix phone number format
    if not phone_number.startswith("+"):
        print(f"Adding country code to: {phone_number}")
        phone_number = "+91" + phone_number

    if not phone_number[1:].isdigit():
        print(f"Invalid number format: {phone_number}. Skipping.")
        continue

    # Compose message
    message_text = (
        f"Dear {name},\n\n"
        f"{row['Message']}\n\n"
        "Hi\n\n"
        "''This is a system-generated message. Please do not reply.''"
    )

    # Encode message for URL
    encoded_message = urllib.parse.quote(message_text)
    url = f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}"

    # Open chat and send message
    driver.get(url)
    try:
        send_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='send']"))
        )
        send_btn.click()
        print(f"Message sent to {name} ({phone_number})")
        time.sleep(5)  # Wait before next message
    except Exception as e:
        print(f"Failed to send message to {name} ({phone_number}): {e}")

print("All messages processed.")
driver.quit()
