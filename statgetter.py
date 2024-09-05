import openpyxl

# Define the sheet name directly in the code
SHEET_NAME = "district"  # Update this to the sheet name you want to use

# Function to load the Excel file
def load_excel_file(file_name):
    try:
        workbook = openpyxl.load_workbook(file_name)
        return workbook
    except FileNotFoundError:
        print(f"Error: The file {file_name} does not exist.")
        exit(1)
    except PermissionError:
        print(f"Error: Permission denied for file {file_name}.")
        exit(1)

# Function to log data into the correct sheet
def log_data_to_excel(workbook, legend, poi, kills, placement, file_name):
    if SHEET_NAME in workbook.sheetnames:
        sheet = workbook[SHEET_NAME]
        # Append the data to the sheet
        sheet.append([legend, poi, kills, placement])
        try:
            workbook.save(file_name)  # Save using the correct file_name variable
            print(f"Data logged in {SHEET_NAME}: {legend}, {poi}, {kills} kills, placed {placement}")
        except PermissionError:
            print(f"Error: Permission denied while saving the file {file_name}.")
    else:
        print(f"Error: {SHEET_NAME} sheet does not exist. Available sheets: {workbook.sheetnames}")

def main():
    file_name = "apexavg.xlsx"  # Use your existing Excel file
    workbook = load_excel_file(file_name)

    # Print existing sheet names for debugging
    print("Current Sheet:\n", SHEET_NAME)
    
        

    while True:
        # Ask the user for inputs
        legend = input("Enter the legend you used: ").strip()
        poi = input("Enter the POI where you landed: ").strip()
        kills = int(input("Enter the number of kills: "))
        placement = int(input("Enter your placement: "))

        # Log the data to the predefined sheet
        log_data_to_excel(workbook, legend, poi, kills, placement, file_name)

        # Ask if the user wants to continue logging
        cont = input("Do you want to log another game? (y/n): ")
        if cont.lower() != 'y':
            break

if __name__ == "__main__":
    main()
