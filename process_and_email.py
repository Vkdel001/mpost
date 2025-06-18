import sys
import pandas as pd
import os
import datetime
import re

raw_text = sys.stdin.read()
lines = raw_text.splitlines()

rows = []
shared_fields = {}
current_postage = {}
current_section = None

def extract_amount(line):
    match = re.search(r'(\d+\.?\d*)', line)  # Improved to handle decimals if needed
    return match.group(1) if match else ''

for line in lines:
    line = line.strip()
    if not line:
        continue

    # New Page (match markdown-style headers like **Page X Report**)
    if re.match(r'\*\*Page\s+\d+\s+Report\*\*', line, re.IGNORECASE):
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

    elif "Registered Postage Section" in line:
        if current_postage:
            row = {**shared_fields, **current_postage}
            rows.append(row)
        current_postage = {"Type": "Registered"}

    elif "Express Postage Section" in line:
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

# Save the last postage block
if current_postage:
    row = {**shared_fields, **current_postage}
    rows.append(row)

# Save to Excel
timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
file_name = f"report_{timestamp}.xlsx"
output_dir = "public"
os.makedirs(output_dir, exist_ok=True)
file_path = os.path.join(output_dir, file_name)

df = pd.DataFrame(rows)
df.to_excel(file_path, index=False)

print(file_name)


 

