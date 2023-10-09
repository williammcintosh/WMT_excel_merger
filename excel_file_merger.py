import openpyxl
import datetime
from datetime import datetime
import calendar
# import pandas as pd

def get_port_waste_tab(input_month, input_year, port_waste_file, verbose=False):
    input_year += 2000
    all_months = list(calendar.month_name)
    first_half = all_months[7:13]
    port_waste_tabs = port_waste_file.sheetnames

    if input_month in first_half:
        tab = f'{input_year}-{input_year+1}'
    else:
        tab = f'{input_year-1}-{input_year}'

    if verbose: print(tab)

    if tab in port_waste_tabs:
        return tab
    
    return None


def excel_merger(verbose=False):
    # Get user input for year and month
    input_year = int(input("Enter the year: "))-2000
    input_month = input("Enter the month: ")

    # Open the port_waste.xlsx file
    try:
        port_waste_file = openpyxl.load_workbook('PORT WASTE COLLECTION  RECY REPORT.xlsx')
    except openpyxl.utils.exceptions.InvalidFileException as e:
        if type(e) == openpyxl.utils.exceptions.InvalidFileException:
            print('''
                ERROR -
                Open 'PORT WASTE COLLECTION  RECY REPORT.xls' in Excel and save as
                'PORT WASTE COLLECTION  RECY REPORT.xlsx' with the 'xlsx' ending,
                then try again. This program doesn't accept 'xls' endings.
            ''')
        else:
            print('An unknown error occurred.')
        return

    # Get the correct tab from the port_waste file
    port_waste_tab = get_port_waste_tab(input_month, input_year, port_waste_file)   
    port_waste_sheet = port_waste_file[port_waste_tab]

    # Open the mapping.xlsx file
    mapping_file = openpyxl.load_workbook('mapping.xlsx')
    mapping_sheet = mapping_file.active

    # Open the master.xlsx file
    master_file = openpyxl.load_workbook('Mastersheets.xlsx')

    # Loop through each row in the mapping sheet
    for map_row in mapping_sheet.iter_rows(min_row=3):

        # Get values for each row from Mapping file
        try:
            site_description = "".join(map_row[0].value.split()).lower()
            material_collected = "".join(map_row[1].value.split()).lower()
            master_tab = map_row[2].value
            master_material_type = "".join(map_row[3].value.split()).lower()
            port_waste_note = None
        except:
            break # EOF, break out!

        # Default values
        master_cell = None
        port_waste_cell = None
        port_waste_value = None

        if verbose: print("\nLocating Row from Port Waste:")
        if verbose: print("\t\t\tFound Val\t'|'\tDesire Val")
        if verbose: print(f"{'-'*60}")
        # Find the row in port_waste.xlsx with both matching site_description and material_collected
        for row in port_waste_sheet.iter_rows(min_row=4):
            if verbose: print(f"\tComparing:\t{row[1].value}\t\t'|'\t{map_row[0].value}")
            if verbose: print(f"\tComparing:\t{row[4].value}\t\t'|'\t{map_row[1].value}")
            if row[1].value is not None and "".join(row[1].value.split()).lower() == site_description and row[4].value is not None and "".join(row[4].value.split()).lower() == material_collected:
                port_waste_row = row[1].row
                break
        else:
            port_waste_row = None

        if verbose: print(f"\nLocating Column from Port Waste:")
        if verbose: print("\t\t\tFound Val\t'|'\tDesire Val")
        if verbose: print(f"{'-'*60}")
        # Find the column in port_waste.xlsx with matching input_month
        for col in port_waste_sheet.iter_cols(min_col=10):
            if verbose: print(f"\tComparing:\t{col[1].value}\t\t'|'\t{input_month}")
            if col[1].value == input_month:
                port_waste_column = col[1].column
                break
        else:
            port_waste_column = None

        # A match was found, store the cell in port_waste.xlsx
        if port_waste_row is not None and port_waste_column is not None:
            port_waste_cell = port_waste_sheet.cell(row=port_waste_row, column=port_waste_column)
            port_waste_value = port_waste_cell.value
            port_waste_note = port_waste_cell.comment.text if port_waste_cell.comment else None
            if verbose: print(f"\nport_waste_row = {port_waste_row}")
            if verbose: print(f"port_waste_column = {port_waste_column}")
            if verbose: print(f"port_waste_cell = {port_waste_cell}")
            if verbose: print(f"port_waste_value = {port_waste_value}")
            if verbose: print(f"port_waste_note = {port_waste_note}")
        else:
            if port_waste_row is None:
                print(f'''
                    Could not find the desired row in the 'PORT WASTE COLLECTION  RECY REPORT.xlsx' file.

                    Compare the spelling of '{map_row[0].value}' under "Site Description" for both the
                    'Mapping.xlsx' file and the 'PORT WASTE COLLECTION  RECY REPORT.xlsx' file.
                    Compare the spelling of '{map_row[1].value}' under "Material Collected" for both the
                    'Mapping.xlsx' file and the 'PORT WASTE COLLECTION  RECY REPORT.xlsx' file. 
                ''')
            if port_waste_column is None:
                print('''
                    WILL!? You done screwed up!
                    Compare the spelling of '{input_month}' with the columns of the
                    'PORT WASTE COLLECTION  RECY REPORT.xlsx' file!
                ''')

        # If the located cell is blank, make it "0"
        port_waste_cell = 0 if port_waste_cell is None else port_waste_cell


        if verbose: print(f"\nLocating Column from Master:")
        if verbose: print("\t\t\tFound Val\t'|'\tDesire Val")
        if verbose: print(f"{'-'*60}")
        # Find the column that matches material type
        master_sheet = master_file[master_tab]
        for col in master_sheet.iter_cols():
            if verbose: print(f"\tComparing:\t{col[2].value}\t\t'|'\t{master_material_type}")
            if col[2].value is not None and "".join(col[2].value.split()).lower() == master_material_type:
                master_column = col[2]
                break
        else:
            master_column = None 

        if verbose: print(f"\nLocating Row from Master:")
        if verbose: print("\t\t\tFound Val\t'|'\tDesire Val")
        if verbose: print(f"{'-'*60}")
        # Find the row that matches the date
        for row in master_sheet.iter_rows(min_row=3):
            input_date = f"{input_month[:3]}-{input_year}"
            master_date = ""
            try:
                date_obj = datetime.strptime(str(row[0].value), "%Y-%m-%d %H:%M:%S")
                year = date_obj.strftime("%y")
                month = date_obj.strftime("%b")
                master_date = f"{month}-{year}"
            except:
                pass

            if verbose: print(f"\tComparing:\t{master_date}\t\t'|'\t"+input_date)
            if master_date == input_date:
                master_row = row[4]
                break
        else:
            master_row = None


        if master_row is not None and master_column is not None:
            master_cell = master_sheet.cell(row=master_row.row, column=master_column.column)
        else:
            if master_column is None and master_row is not None:
                print(f"""
                    Could not find the desired row in the 'Mastersheet.xlsx' file.\n
                    Compare the spelling of '{map_row[2].value}' under "Tab" for both the
                    'Mapping.xlsx' file and the 'Mastersheet.xlsx' file.
                    Compare the spelling of '{map_row[3].value}' under "Type" for both the
                    'Mapping.xlsx' file and the 'Mastersheet.xlsx' file.
                """)
            elif master_row is None and master_column is not None:
                print("""
                    WILL!? You done screwed up!
                    Compare the spelling of '{input_month}' with the rows of the
                    'Mastersheet.xlsx' file!
                """)

        # Update the formula in the master cell
        if master_cell is not None and port_waste_value is not None:
            
            formula = f"{master_cell.value}+{port_waste_value}" if master_cell.value is not None else f"{port_waste_value}" 
            master_cell.value = formula
            if port_waste_note is not None:
                master_cell.comment = openpyxl.comments.Comment(port_waste_note, "Author")

        if verbose: print(f"\nport_waste_cell = {port_waste_cell}")
        if verbose: print(f"master_cell = {master_cell}")
        if verbose: print("\n")

    # Save the changes to master.xlsx
    current_datetime = datetime.now()
    formatted_datetime = current_datetime.strftime("%Y-%m-%d_%H-%M-%S")
    new_file_name = f"master_{formatted_datetime}.xlsx"
    master_file.save(new_file_name)

if __name__ == "__main__":
    verbose = input("Verbose? ([y]/n): ")
    verbose = True if verbose.lower() == 'y' else False
    excel_merger(verbose)
