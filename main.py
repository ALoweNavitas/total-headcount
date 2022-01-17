import pandas as pd
import xlwings as xw
import os

dir = os.getcwd()

input = []
output = []

# Locate the export file
for file in os.listdir(os.path.join(dir, 'export')):
    if file.endswith('.xlsx'):
        fp = os.path.basename(file)
        input.append(os.path.join(dir, "export", fp).replace("\\", "/"))

# Locate the output file
for file in os.listdir(os.path.join(dir, 'output')):
    if file.endswith('.xlsx'):
        fp = os.path.basename(file)
        output.append(os.path.join(dir, 'output', fp).replace("\\", "/"))

# Read in the export file
df = pd.read_excel(input[0])

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
workbook = xw.Book(output[0], password='password')
ws = workbook.sheets['head count']
ws.range('A4').options(index=False, header=False).value = df

