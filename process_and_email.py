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

def extract_amount(line):
    match = re.search(r'(\d+)', line)
    return match.group(1) if match else ''

current_section = None

for line in lines:
    line = line.strip()
    if not line:
        continue

    if line.lower().startswith("**page"):
        # If there's a previous postage block, save it before moving to new page
        if current_postage:
            row = {**shared_fields, **current_postage}
            rows.append(row)
            current_postage = {}

        shared_fields = {}
        shared_fields["Page"] = extract_amount(line)

    elif "Date" in line:
        shared_fields["Date"] = extract_amount(line)

    elif "Invoice Number" in line:
        shared_fields["Invoice No"] = extract_amount(line)

    elif "MIN Number" in line:
        shared_fields["MIN No"] = extract_amount(line)

    elif "Department" in line:
        shared_fields["Department"] = line.split(":", 1)[-1].strip()

    elif "Registered Postage" in line:
        # Save previous section if exists
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

# Add last row if not yet saved
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
