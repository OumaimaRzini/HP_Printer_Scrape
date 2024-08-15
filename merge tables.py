import pandas as pd

# Read the tables from Excel files specifying the sheets
workcenter_table = pd.read_excel("Printers inventory.xlsx", sheet_name="WorkCenter")
printer_usage_table = pd.read_excel("Printer_Metrics.xlsx", sheet_name="Printer Daily Usage")

# Merge the tables based on IP address
merged_table = pd.merge(printer_usage_table, workcenter_table, left_on='IP Address', right_on='IP Address', how='left')

# Drop unnecessary columns from the merged table
merged_table.drop(['WorkCenter', 'Poste', 'Line ID', 'LRS name'], axis=1, inplace=True)

# Rename the WorkCenter ID column
merged_table.rename(columns={'WorkCenter ID': 'WorkCenter ID'}, inplace=True)

# Save the merged table to the same Excel file but in a different sheet
with pd.ExcelWriter("Printer_Metrics.xlsx", engine='openpyxl', mode='a') as writer:
    merged_table.to_excel(writer, index=False, sheet_name="Workcenter Printers")

