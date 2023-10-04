import openpyxl
import datetime
from datetime import datetime

# Get user input for year and month
input_year = int(input("Enter the year: "))-2000
input_month = input("Enter the month: ")

# Open the port_waste.xlsx file
port_waste_file = openpyxl.load_workbook('port_waste_example.xlsx')
port_waste_sheet = port_waste_file['2023-2024']  # Update the sheet name accordingly

# Open the mapping.xlsx file
mapping_file = openpyxl.load_workbook('mapping.xlsx')
mapping_sheet = mapping_file.active

# Open the master.xlsx file
master_file = openpyxl.load_workbook('master_example.xlsx')

# Loop through each row in the mapping sheet
for map_row in mapping_sheet.iter_rows(min_row=3):
    site_description = map_row[0].value
    material_collected = map_row[1].value
    master_tab = map_row[2].value
    master_material_type = map_row[3].value
    master_cell = None
    port_waste_cell = None

    print(f"Locating Row from Port Waste:")
    # Find the corresponding row in port_waste.xlsx
    for row in port_waste_sheet.iter_rows(min_row=5):
        print(f"\tComparing:\t{row[1].value}\t{site_description}")
        print(f"\tComparing:\t{row[4].value}\t{material_collected}")
        if row[1].value == site_description and row[4].value == material_collected:
            port_waste_row = row[1]
            break
    else:
        port_waste_row = None

    print(f"Locating Column from Port Waste:")
    # Find the corresponding column in port_waste.xlsx
    for col in port_waste_sheet.iter_cols(min_col=10):
        print(f"\tComparing:\t{col[1].value}\t{input_month}")
        if col[1].value == input_month:
            port_waste_column = col[1]
            break
    else:
        port_waste_column = None


    if port_waste_row is not None and port_waste_column is not None:
        port_waste_cell = port_waste_sheet.cell(row=port_waste_row.row, column=port_waste_column.column)


    print(f"Locating Column from Master:")
    # Find the column that matches material type
    master_sheet = master_file[master_tab]
    for col in master_sheet.iter_cols():
        print(f"\tComparing:\t{col[2].value}\t{master_material_type}")
        if col[2].value == master_material_type:
            master_column = col[2]
            break
    else:
        master_column = None

    print(f"Locating Row from Master:")
    # Find the row that matches the date
    for row in master_sheet.iter_rows(min_row=3):
        input_date = f"{input_month}-{input_year}"
        master_date = ""
        try:
            date_obj = datetime.strptime(str(row[0].value), "%Y-%m-%d %H:%M:%S")
            year = date_obj.strftime("%y")
            month = date_obj.strftime("%b")
            master_date = f"{month}-{year}"
        except ValueError:
            print("\tInvalid date format")
        print(f"\tComparing:\t{master_date}\t"+input_date)
        if master_date == input_date:
            master_row = row[4]
            break
    else:
        master_row = None

    if master_row is not None and master_column is not None:
        master_cell = master_sheet.cell(row=master_row.row, column=master_column.column)

    # Update the formula in the master cell
    if master_cell is not None and port_waste_cell is not None:
        formula = f"={master_cell.value}+{port_waste_cell.value}"
        master_cell.value = formula

    print("")
    print(f"={port_waste_cell.value}")
    print(f"={master_cell}")

# Save the changes to master.xlsx
current_datetime = datetime.now()
formatted_datetime = current_datetime.strftime("%Y-%m-%d_%H-%M-%S")
new_file_name = f"master_{formatted_datetime}.xlsx"
# master_file.save(new_file_name)
