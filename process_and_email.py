import sys
import pandas as pd
import os
import datetime
import re

raw_text = sys.stdin.read()
lines = raw_text.splitlines()

data = []
current_page = None
row = {}

def extract_amount(line):
    match = re.search(r'(\d+)', line)
    return match.group(1) if match else ''

for line in lines:
    line = line.strip()
    if not line:
        continue

    if line.lower().startswith("**page"):
        if row:
            data.append(row)
        row = {}
        current_page = extract_amount(line)
        row["Page"] = current_page

    elif "Date" in line:
        row["Date"] = extract_amount(line)

    elif "Invoice Number" in line:
        row["Invoice No"] = extract_amount(line)

    elif "MIN Number" in line:
        row["MIN No"] = extract_amount(line)

    elif "Department" in line:
        row["Department"] = line.split(":", 1)[-1].strip()

    elif "Registered Postage" in line:
        current_section = "Registered"

    elif "Express Postage" in line:
        current_section = "Express"

    elif "Quantity" in line:
        row[f"{current_section} Quantity"] = extract_amount(line)

    elif "Rate" in line:
        row[f"{current_section} Rate"] = line.split(":", 1)[-1].strip()

    elif "Total Amount Postage" in line:
        row[f"{current_section} Postage"] = extract_amount(line)

    elif "Total Amount Fee" in line:
        fee = extract_amount(line)
        row[f"{current_section} Fee"] = fee if fee else "0"

    elif "Total Amount Payable" in line:
        row[f"{current_section} Payable"] = extract_amount(line)

    elif "Total Number of Letters Posted" in line:
        row["Total Letters"] = extract_amount(line)

# Add the last page
if row:
    data.append(row)

# Save to Excel
timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
file_name = f"report_{timestamp}.xlsx"
output_dir = "public"
os.makedirs(output_dir, exist_ok=True)
file_path = os.path.join(output_dir, file_name)

df = pd.DataFrame(data)
df.to_excel(file_path, index=False)

print(file_name)
