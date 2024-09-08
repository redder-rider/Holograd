import pandas as pd

data = """
Name: Usare Ashvini Kaduba 
Email ID:  usareashvini01@gmail.com
Contact : 8459833953
Department : HRM
Date of joining : 1 August 2024
Team leader: Pratham Tripude
"""

lines = data.strip().split("\n\n")
records = []

for line in lines:
    record = {}
    for field in line.split("\n"):
        parts = field.split(": ", 1)
        if len(parts) == 2:
            key, value = parts
            key = key.replace(" ", "")
            record[key] = value
    records.append(record)

df = pd.DataFrame(records)

excel_file_path = r"c:\Users\daksh\OneDrive\Desktop\employees.xlsx"

df.to_excel(excel_file_path, index=False, engine='openpyxl')

print("Excel file has been created successfully.")
