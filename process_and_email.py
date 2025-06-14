import sys
import pandas as pd
import re
import os
import datetime

raw_text = sys.stdin.read()
raw_text = raw_text.encode('utf-8').decode('utf-8')  # ensure utf-8

# Split pages using pattern like "**Page 1:**"
pages = re.split(r'\*\*Page \d+:?\*\*', raw_text)
data = []

for i, page in enumerate(pages[1:], start=1):  # Skip first split (header)
    row = {'Page': i}

    def extract(pattern):
        match = re.search(pattern, page)
        return match.group(1).strip() if match else ''

    # Updated flexible patterns (allowing optional ** and extra spaces)
    row['Invoice No'] = extract(r'\*\*Invoice No:\*\*\s*(\d+)')
    row['Date'] = extract(r'\*\*Date:\*\*\s*([\d/]+)')
    row['Ministry'] = extract(r'\*\*Ministry:\*\*\s*(.*)')
    row['Department'] = extract(r'\*\*Department:\*\*\s*(.*)')

    # Handle Registered Letters block (optional multiple @ Rate)
    registered_block = re.findall(r'\*\*(\d+)\s*@\s*(\d+)\s*Rs:\*\*\s*(\d+)', page)
    if registered_block:
        rates_summary = ', '.join([f'{q} @ {r} = {a}' for q, r, a in registered_block])
    else:
        quantity = extract(r'\*\*Quantity:\*\*\s*(\d+)')
        rate = extract(r'\*\*Rate:\*\*\s*(\d+)')
        rates_summary = f'{quantity} @ {rate}' if quantity and rate else ''
    row['Registered Rate(s)'] = rates_summary

    # Updated Fee, Postage, and Total extraction patterns
    row['Total Amt Postage (Rs)'] = extract(r'\*\*Total Amt Postage \(Rs\):\*\*\s*(\d+)')
    row['Total Amt Fee (Rs)'] = extract(r'\*\*Total Amt Fee \(Rs\):\*\*\s*(\d+)')
    row['Total Amt (Rs)'] = extract(r'\*\*Total Amt \(Rs\):\*\*\s*(\d+)')
    row['Total Letters Posted'] = extract(r'\*\*Total Number of Letters Posted:\*\*\s*(\d+)')
    row['Total Amount Payable (Rs)'] = extract(r'\*\*Total Amount Payable \(Rs\):\*\*\s*(\d+)')

    data.append(row)

# Create Excel file
timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
file_name = f'report_{timestamp}.xlsx'
output_dir = 'public'
os.makedirs(output_dir, exist_ok=True)
file_path = os.path.join(output_dir, file_name)

df = pd.DataFrame(data)
df.to_excel(file_path, index=False)

# Return filename to Node.js
print(file_name)
