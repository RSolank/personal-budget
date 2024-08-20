import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import random
from datetime import datetime, timedelta

# Define the headers
headers = ["Tran Date", "Chq No", "Particulars", "Category", "Sub-Category", "Debit", "Credit", "Balance", "Init. Br"]

# Generate random data
def generate_random_data(num_rows):
    data = []
    start_date = datetime.now() - timedelta(days=365)
    
    for _ in range(num_rows):
        date = start_date + timedelta(days=random.randint(0, 365))
        chq_no = f"{random.randint(1000, 9999)}"
        particulars = f"Detail {random.randint(1, 100)}"
        category = f"Category {random.randint(1, 10)}"
        sub_category = f"Sub-Category {random.randint(1, 5)}"
        debit = round(random.uniform(0, 1000), 2)
        credit = round(random.uniform(0, 1000), 2)
        balance = debit - credit
        init_br = round(random.uniform(0, 5000), 2)
        data.append([date, chq_no, particulars, category, sub_category, debit, credit, balance, init_br])
    
    return data

# Create a new workbook and select the active worksheet
wb = openpyxl.Workbook()
ws = wb.active

# Write the headers to the first row
ws.append(headers)

# Generate random data and write it to the worksheet
data = generate_random_data(100)  # Adjust the number of rows as needed
for row in data:
    ws.append(row)

# Save the workbook to a file
file_name = "dummy_transactions.xlsx"
wb.save(file_name)

print(f"Dummy Excel file '{file_name}' created successfully.")
