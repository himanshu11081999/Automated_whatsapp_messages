from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import urllib.parse
import os

# Load Excel File
file_path = "E:\\CMC\\Demo.xlsx"
df = pd.read_excel(file_path, engine='openpyxl')

# Check necessary columns
if "Phone" not in df.columns or "Message" not in df.columns or "Name" not in df.columns:
    raise ValueError("The Excel file must contain 'Phone', 'Name', and 'Message' columns!")

# Set ChromeDriver path
chromedriver_path = "E:\\CMC\\Miscellaneous\\Java\\cd\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe"  # Path to your ChromeDriver

# Path to the image file
image_path = "E:\\CMC\\RG marathon.png"  # Update with your image path

if not os.path.exists(image_path):
    raise FileNotFoundError(f"The image file was not found at: {image_path}")

# Initialize Chrome with the updated Service class
service = Service(chromedriver_path)
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

# Open WhatsApp Web
driver.get("https://web.whatsapp.com")
print("Scan the QR code in the browser.")

# Wait for WhatsApp Web login to complete
WebDriverWait(driver, 60).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#side > div.x1ky8ojb.x78zum5.x1q0g3np.x1a02dak.x2lah0s.x3pnbk8.xfex06f.xeuugli.x2lwn1j.x1nn3v0j.x1ykpatu.x1swvt13.x1pi30zi"))
)

# Preprocess phone numbers
for index, row in df.iterrows():
    phone_number = str(row["Phone"]).strip()
    name = str(row["Name"]).strip()  # Extract name from the Excel sheet

    # Ensure the phone number starts with '+'
    if not phone_number.startswith("+"):
        print(f"Invalid format detected: {phone_number}. Adding '+91' to the beginning.")
        phone_number = "+91" + phone_number

    # Ensure the number contains only digits (after '+')
    if not phone_number[1:].isdigit():
        print(f"Invalid phone number: {phone_number}. Skipping.")
        continue

    # Create the personalized message
    message_text = (
        "Greetings From CMC Ludhiana,\n"
    )

    # Encode the message text for the URL
    encoded_message = urllib.parse.quote(message_text)

    # Generate WhatsApp Web URL
    url = f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}"
    driver.get(url)

    try:

        # Attach and send the image
        attach_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='plus']"))
        )
        attach_button.click()

        # Upload the image file
        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
        )
        file_input.send_keys(os.path.abspath(image_path))
        time.sleep(3)  # Wait for the image preview to load

        # Click the send button for the image
        send_image_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='send']"))
        )
        send_image_button.click()
        print(f"Image sent to {name} ({phone_number})")

        time.sleep(1)  # Delay before processing the next message
        
        # Wait for the send button to appear and click
        send_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='send']"))
        )
        send_button.click()
        print(f"Message sent to {name} ({phone_number})")

        time.sleep(1)  # Delay before sending the image

        

    except Exception as e:
        print(f"Failed to send message or image to {name} ({phone_number}). Error: {e}")
        continue

print("All messages and images sent successfully!")
driver.quit()
