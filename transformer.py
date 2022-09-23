from openpyxl import Workbook, load_workbook
from os import listdir

columns_map:dict[str, str] = {}
wb:Workbook

def convert_sheet_to_csv(sheet_name:str):
    print(f"Converting sheet \"{sheet_name}\".")
    sheet = wb[sheet_name]

if __name__ == "__main__":
    #try:
        config:str = open("./.config")
        for line in config:
            if line.startswith("//") or len(line) == 0 or line.isspace():
                continue

            line = line.strip()
            values:list[str] = line.split("=")
            columns_map[values[0]] = values[1]
            
            if len(values[1]) == 0 or values[1].isspace():
                print(f"Warning: The column \"{values[0]}\" is unset.")

        xl_files:list[str] = [file for file in listdir(".") if ".xlsx" in file]

        if len(xl_files) == 0:
            raise Exception("No .xlsx files found in this directory.")
            
        xl_file_name:str = xl_files[0]
        wb = load_workbook(xl_file_name)
        print(f"Converting \"{xl_file_name}\".")

        for sheet in wb.sheetnames:
            convert_sheet_to_csv(sheet)
    
    # except Exception as e:
    #     input(f"\nError: {str(e)}\nPress ENTER to exit.")