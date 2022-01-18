import pandas as pd
import xlwings as xw
import os
import time

dir = os.getcwd()
os.chdir(dir)

import_file = []
output_file = []

# Locate the export file
for file in os.listdir(os.path.join(dir, 'export')):
    if file.endswith('.xlsx'):
        fp = os.path.basename(file)
        import_file.append(os.path.join(dir, 'export', fp).replace("\\", "/"))

# Locate the output file
for file in os.listdir(os.path.join(dir, 'output')):
    if file.endswith('.xlsx'):
        fp = os.path.basename(file)
        output_file.append(os.path.join(dir, 'output', fp).replace("\\", "/"))

# Read in the export file
df = pd.read_excel(import_file[0])

# Filter the dataframe to only have UPE records
upe_filter = df['Contract - Team'] == 'UPE'
df = df.loc[upe_filter, :]

# Reshape to requirements of the template Excel
df = df[[
    "Contract - Start Date",
    "Contract - Job Type",
    "Contract - Team", 
    "Contract - Business Unit",
    "Personal - Gender",
    "Personal - Date of Birth",
    "Contract - Contract Type"
]]

# Export to Excel
workbook = xw.Book(output_file[0], password='password')
ws = workbook.sheets['head count']
ws.range('A4').options(index=False, header=False).value = df

# time.sleep(5)
workbook.save()

def countdown(time_sec):
    while time_sec:
        mins, sec = divmod(time_sec, 60)
        timeformat = '{:02d}:{:02d}'.format(mins, sec)
        print(timeformat, end='\r')
        time.sleep(1)
        time_sec -= 1

countdown(5)