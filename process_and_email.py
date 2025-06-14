import sys
import pandas as pd
import re
import os

raw_text = sys.stdin.read()
pages = re.split(r'\*\*Page \d+\*\*', raw_text)
data = []

for i, page in enumerate(pages[1:], start=1):
    match_dict = {
        'Invoice No': re.search(r'Invoice No:\**\s*(\d+)', page),
        'Date': re.search(r'Date:\**\s*([\d/]+)', page),
        'Ministry/Department': re.search(r'(?:Ministry|Department):\**\s*(.+)', page),
        'Total Letters': re.search(r'Total No\. of Letters Posted:\**\s*(\d+)', page),
        'Registered Rate(s)': re.search(r'Registered Rate[s]*:\**\s*(.*?)\n(?:Charges:|Rupees:)', page, re.DOTALL),
        'Rupees': re.search(r'Rupees:\**\s*([\d+ ]+[\d+]*)', page),
        'Fee Amount': re.search(r'Fee Amount:\**\s*(\d+)', page),
        'Total Amount Payable': re.search(r'Total Amount Payable:\**\s*(\d+)', page)
    }

    row = {'Page': i}
    for k, v in match_dict.items():
        try:
            value = v.group(1).strip().replace('\n', ', ') if v else ''
        except Exception as e:
            value = ''
        row[k] = value
    data.append(row)

df = pd.DataFrame(data)
output_dir = 'public'
os.makedirs(output_dir, exist_ok=True)
file_path = os.path.join(output_dir, 'report.xlsx')
df.to_excel(file_path, index=False)

print(f"✅ Excel report created and saved to: {file_path}")
