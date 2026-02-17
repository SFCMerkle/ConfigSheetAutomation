
import openpyxl as pxl

def main():
    file_path = input("Enter the path to your Excel file: ")
    #file_path = "C:/Users/cmerkle/OneDrive - IDEX Corporation/Desktop/Configsheet_Project/CS-BRAC055U007.xlsx"
    
    #file_path = "N:/SFC Engineering/Switzerland/FLOW/1 Configurationsheet/XLS/02_Restrictor/RxAA/5.5 mm/RR-S/Released/CS-RRAA055S7758.xlsx"
    file_path = file_path.replace("\\", "/")  # Replace backslashes with forward slashes


    Config_file = pxl.load_workbook(file_path, data_only=True)
    # Get the last sheet in the workbook 
    Config_sheet = Config_file.worksheets[-1]
    print(Config_sheet.title)



    positions = []
    for row in Config_sheet.iter_rows():
        for cell in row:
            if cell.value and "Title" in str(cell.value):
                positions.append((cell.row, cell.column))

    if not positions:
        print("No cells with 'Title' found.")
        return

    #print("Cells with 'Title':", positions)
    
    partcount = 1       
    for pos in positions:
        
        row, col = pos
        value_cell = Config_sheet.cell(row=row, column=col+1)
        
        while not value_cell.value and value_cell.column <= Config_sheet.max_column:
            value_cell = Config_sheet.cell(row=row, column=value_cell.column+1)
        print(f" Part Title {partcount}: {value_cell.value}")
        partcount += 1


if __name__ == "__main__":
    main()
    

