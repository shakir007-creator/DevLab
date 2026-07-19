from pathlib import Path
import xlwings as xw

def clean_data():

    # Load the workbook and sheet using xlwings
    file_path = Path.cwd() / "Excel_File" / "Latest_B2B.xlsx"  # Replace with your file path

    print("\nCleaning Data...\n")

    del_crq = 0
    del_ckt = 0

    try:
        # Start Excel in the background (visible=False) to prevent UI from popping up
        with xw.App(visible=False) as app:
            try:
                wb = app.books.open(file_path)  # Open the workbook
            except Exception as e:
                print(f"\nError opening the workbook: {e}\n")
                raise  # Re-raise the exception if opening the workbook fails

            try:
                sheet = wb.sheets['Lookup-Sheet']  # Access the sheet by name
            except KeyError as e:
                print(f"\nError: Sheet 'Lookup-Sheet' not found in the workbook: {e}\n")
                raise  # Re-raise the exception if the sheet is missing

            # Find the last non-empty row from the bottom in column D
            try:
                last_row = sheet.range('D' + str(sheet.cells.last_cell.row)).end('up').row  # Find the last row in column D
            except Exception as e:
                print(f"\nError finding the last row in column D: {e}\n")
                raise  # Re-raise the exception if finding the last row fails

            # Iterate over rows from bottom to top
            for row in range(last_row, 1, -1):  # Start from last_row and go upwards
                try:
                    cell_d_value = sheet.range(f'D{row}').value  # Read the value in column D
                    cell_e_value = sheet.range(f'E{row}').value  # Read the value in column E

                    # Check if "No Data Found" is in the value of column E
                    if cell_e_value == "No Data Found":
                        # Check if the cell below is empty and the cell above contains "CRQ"
                        if (row < last_row and not sheet.range(f'D{row+1}').value) and 'CRQ' in str(sheet.range(f'D{row-1}').value):  # Below cell is empty and CRQ in Above cell
                            # Delete the current row and the row above it
                            print(f"\n{sheet.range(f'D{row-1}').value} Deleted")
                            print(f"\nCKT ID: {int(cell_d_value)} Deleted\n")
                            sheet.range(f'{row-1}:{row}').api.Delete()  # Delete two rows (current row and above row)
                            last_row -= 2  # Adjust last row index due to row deletion
                            del_crq += 1
                            del_ckt +=1
                        else:
                            # If the above cell doesn't contain "CRQ", just delete the current row
                            print(f"\nCKT ID: {int(cell_d_value)} Deleted\n")
                            sheet.range(f'{row}:{row}').api.Delete()  # Delete the current row
                            last_row -= 1  # Adjust last row index due to row deletion
                            del_ckt +=1
                except Exception as e:
                    print(f"\nError processing row {row}: {e}\n")
                    continue  # Skip the current row and continue with the next row

            # Save the changes to the workbook
            try:
                wb.save(file_path)  # Save the modified file
                if del_crq == 0 and del_ckt == 0:
                    print("\nNO DATA TO DELETE.\n")
                else:
                    print("\nData Cleaned Successfully.\n")
                    print(f"\n{del_crq} CRQs Deleted.\n")
                    print(f"\n{del_ckt} CKTs Deleted.\n")
            except Exception as e:
                print(f"\nError saving the workbook: {e}\n")
                raise  # Re-raise the exception if saving fails

    except FileNotFoundError as e:
        print(f"\nError: The file {file_path} was not found. Please check the file path.\n")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}\n")

    print("\n======================================================================\n\n")

if __name__ == "__main__":
    clean_data()
