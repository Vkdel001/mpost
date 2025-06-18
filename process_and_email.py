import sys
import pandas as pd
import os
import datetime
import re
import requests

# Read raw text input from stdin
raw_text = sys.stdin.read()
lines = raw_text.splitlines()

rows = []
shared_fields = {}
current_postage = {}

def extract_amount(line):
    match = re.search(r'(\d+)', line)
    return match.group(1) if match else ''

for line in lines:
    line = line.strip()
    if not line:
        continue

    # New Page
    if re.match(r'[#\*\s]*page\s+\d+', line.lower()):
        if current_postage:
            row = {**shared_fields, **current_postage}
            rows.append(row)
            current_postage = {}
        shared_fields = {}
        shared_fields["Page"] = extract_amount(line)
        continue

    # Detect and assign shared fields
    if "Date" in line and "Total" not in line:
        shared_fields["Date"] = line.split(":", 1)[-1].strip()

    elif "Invoice Number" in line:
        shared_fields["Invoice No"] = extract_amount(line)

    elif "MIN Number" in line:
        shared_fields["MIN No"] = extract_amount(line)

    elif "Department Name" in line:
        shared_fields["Department"] = line.split(":", 1)[-1].strip()

    elif "Registered Postage" in line:
        if current_postage:
            row = {**shared_fields, **current_postage}
            rows.append(row)
        current_postage = {"Type": "Registered"}

    elif "Express Postage" in line:
        if current_postage:
            row = {**shared_fields, **current_postage}
            rows.append(row)
        current_postage = {"Type": "Express"}

    elif "Quantity" in line:
        current_postage["Quantity"] = extract_amount(line)

    elif "Rate" in line:
        current_postage["Rate"] = line.split(":", 1)[-1].strip()

    elif "Total Amount Postage" in line:
        current_postage["Postage"] = extract_amount(line)

    elif "Total Amount Fee" in line:
        fee = extract_amount(line)
        current_postage["Fee"] = fee if fee else "0"

    elif "Total Amount Payable" in line:
        current_postage["Payable"] = extract_amount(line)

    elif "Total Number of Letters Posted" in line:
        shared_fields["Total Letters"] = extract_amount(line)

# Save last postage block
if current_postage:
    row = {**shared_fields, **current_postage}
    rows.append(row)

# Save to Excel file
timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
file_name = f"report_{timestamp}.xlsx"
output_dir = "public"
os.makedirs(output_dir, exist_ok=True)
file_path = os.path.join(output_dir, file_name)

df = pd.DataFrame(rows)
df.to_excel(file_path, index=False)

print(f"📁 Excel saved as: {file_name}")

# ✅ Upload to GoFile.io using your account token
token = "HSl6SrcLB3j9dCCVcSHagTRD2jFaKkJO"
upload_url = "https://api.gofile.io/uploadFile"

try:
    with open(file_path, "rb") as f:
        files = {"file": f}
        data = {"token": token}
        response = requests.post(upload_url, files=files, data=data)

    result = response.json()
    print("🔍 API Response:", result)

    if result["status"] == "ok":
        download_link = result["data"]["downloadPage"]
        print("✅ File uploaded to GoFile.io")
        print("🔗 Download Link:", download_link)
    else:
        print("❌ Upload failed:", result.get("status"))

except Exception as e:
    print("❌ Error during upload:", str(e))
