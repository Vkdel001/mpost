import sys
import pandas as pd
import re
import os
import datetime

# Read all input from stdin
raw_text = sys.stdin.read()

# Normalize input: remove double asterisks (**) but keep content
clean_text = re.sub(r'\*\*', '', raw_text)

# Flexible split by "Page <number>" ignoring special chars like ** or :
# This handles variations like **Page 1:**, Page 1, etc.
pages = re.split(r'Page\s+\d+', clean_text, flags=re.IGNORECASE)

data = []

for i, page in enumerate(pages[1:], start=1):
    # Patterns allow optional spaces, colons, and ignore case
    invoice_no = re.search(r'Invoice\s*No[:\s]*([\d]+)', page, re.IGNORECASE)
    date = re.search(r'Date[:\s]*([\d/]+)', page, re.IGNORECASE)
    ministry = re.search(r'Ministry[:\s]*([\w\s]+)', page, re.IGNORECASE)
    department = re.search(r'Department[:\s]*([\w\s]*)', page, re.IGNORECASE)
    total_letters = re.search(r'Total\s*(No\.?|Number)?\s*of\s*Letters\s*Posted[:\s]*([\d]+)', page, re.IGNORECASE)
    # For registered letters, sum quantities and rates if possible
    # We'll just capture whole Registered Letters section text for now
    registered_letters_section = re.search(r'Registered\s*Letters[:\s]*([\s\S]*?)(?:Total|$)', page, re.IGNORECASE)
    fee_amount = re.search(r'Fee\s*Amount[:\s]*([\d]+)', page, re.IGNORECASE)
    total_payable = re.search(r'Total\s*Amount\s*Payable[:\s]*([\d]+)', page, re.IGNORECASE)

    row = {
        'Page': i,
        'Invoice No': invoice_no.group(1).strip() if invoice_no else '',
        'Date': date.group(1).strip() if date else '',
        'Ministry': ministry.group(1).strip() if ministry else '',
        'Department': department.group(1).strip() if department else '',
        'Total Letters Posted': total_letters.group(2).strip() if total_letters else '',
        'Registered Letters Details': registered_letters_section.group(1).strip().replace('\n', ', ') if registered_letters_section else '',
        'Fee Amount': fee_amount.group(1).strip() if fee_amount else '',
        'Total Amount Payable': total_payable.group(1).strip() if total_payable else ''
    }

    data.append(row)

# Create filename with timestamp
timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
file_name = f'report_{timestamp}.xlsx'

output_dir = 'public'
os.makedirs(output_dir, exist_ok=True)
file_path = os.path.join(output_dir, file_name)

# Save dataframe to Excel
df = pd.DataFrame(data)
df.to_excel(file_path, index=False)

# Output just the filename for your Node.js to pick up
print(file_name)
