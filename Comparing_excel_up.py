import os
import openpyxl
from openpyxl import load_workbook


def format_upExcel(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location):
    
    # List all files in each folder
    last_month_files = os.listdir(last_month_file_location)
    formatted_files = [x for x in os.listdir(file_path_to_output + "/FORMATTED")]

    def get_headers(file_path):     # Function to get headers from an Excel file
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        return [cell.value for cell in sheet[1]]

    # Compare each pair of files
    for last_month_file in last_month_files:
        last_month_file1=last_month_file.split('-')[0]
        for formatted_file in formatted_files:
            formatted_file1=formatted_file.split('-')[0]
            if last_month_file1 == formatted_file1:
                 last_month_path = os.path.join(last_month_file_location, last_month_file)
                 formatted_path = os.path.join(file_path_to_output + "/FORMATTED", formatted_file)
                 last_month_headers = get_headers(last_month_path)
                 formatted_headers = get_headers(formatted_path )
                 new_columns = set(formatted_headers) - set(last_month_headers)     # Find new columns added in the formatted file
                 print(f"\n Comparison result for '{last_month_file}' and '{formatted_file}':")
                 if new_columns:
                     
                    print("\n New columns added in formatted file:")
                    for column in new_columns:
                        
                       print(column)
                 else:
                  print("\n No new columns added in formatted file.")
                  
                 break
            else:
                 continue
             
#if __name__ == "__main__":
#    main(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location)
#format_upExcel(file_path_to_output, brands_mapping_location, map_api_location, map_location, last_month_file_location)