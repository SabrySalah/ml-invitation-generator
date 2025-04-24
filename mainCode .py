import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import pywhatkit as kit
import os
import time

# === Configuration ===
EXCEL_PATH = 'customers.xlsx'
TEMPLATE_PATH = 'invitation_template.jpg'
OUTPUT_DIR = 'output_invitations'
FONT_PATH = 'arial.ttf'
FONT_SIZE = 36

# Setup
os.makedirs(OUTPUT_DIR, exist_ok=True)
df = pd.read_excel(EXCEL_PATH)
font = ImageFont.truetype(FONT_PATH, FONT_SIZE)

# === Loop through each row ===
for index, row in df.iterrows():
    name = str(row['Name'])
    company = str(row['Company'])
    phone = str(row['Phone'])

    # === Create personalized invitation ===
    img = Image.open(TEMPLATE_PATH).convert("RGB")
    draw = ImageDraw.Draw(img)
    
    name_position = (100, 300)
    company_position = (100, 370)

    draw.text(name_position, name, fill="black", font=font)
    draw.text(company_position, company, fill="black", font=font)

    image_filename = f"{name}_{company}.jpg"
    image_path = os.path.join(OUTPUT_DIR, image_filename)
    img.save(image_path)

    print(f"Generated invitation for {name}")

    # === Send WhatsApp message ===
    try:
        full_phone = f'+2{phone}'  # Egypt country code, adjust if needed
        message = f"Hi {name}, here’s your invitation for the event!"
        kit.sendwhats_image(
            phone_no=full_phone,
            img_path=image_path,
            caption=message,
            wait_time=10
        )
        print(f"Sent invitation to {phone}")
    except Exception as e:
        print(f"Failed to send to {phone}: {e}")

    time.sleep(5)  # Prevents sending too fast
