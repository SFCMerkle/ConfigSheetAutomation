
import pandas as pd
import openpyxl as pxl


# Path to your Excel file
file_path = "C:/Users/cmerkle/OneDrive - IDEX Corporation/Desktop/Configsheet_Project/CS-BRAC055U007.xlsx"
#print(f"Filename: {file_path}")

# Read the Excel file
Config_file = pxl.load_workbook(file_path, data_only=True)
last_sheet = Config_file[Config_file.sheetnames[-1]]

# Display the second row of the last sheet

print(f"S2, Version: {last_sheet['S2'].value}")


