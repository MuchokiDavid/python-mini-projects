# This project describe how to use python to manipulate data stored in your google sheets
# Install gpread and oauth2client using pip
# Create a credentials.json file in your project directory and paste your credentials there
# Make sure you have enabled the google sheets api and google drive api in your google console account
# Run the code and enjoy

# Importing required libraries
import gspread

from oauth2client.service_account import ServiceAccountCredentials

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json",scope)
client = gspread.authorize(creds)

# Read google sheet file
sheet = client.open("Python test").sheet1


#Getting data from the sheet
data = sheet.get_all_records()

# print(data)

#Get the specific row, column and cell from the sheet

row = sheet.row_values(3)
col = sheet.col_values(3)
cell = sheet.cell(1,2).value

# print(cell)

#Inserting data in your sheet 

# insertRow = [11,"Zayn","Malik",12,400]
# sheet.insert_row(insertRow,12)
# print("The row has been added")

# # delete data
# sheet.delete_rows(10)
# print("The row has been deleted")

# query data
query = sheet.find("Mel")
print(query.col, query.row)


# Methods and Operations

# Spreadsheet Operations
# open(title): Opens a spreadsheet by its title.
# open_by_key(key): Opens a spreadsheet using its unique key (found in the URL of the spreadsheet).
# open_by_url(url): Opens a spreadsheet using its URL.
# create(title): Creates a new spreadsheet with the specified title.
# copy(file_id): Creates a copy of the spreadsheet identified by the given file ID.
# list_spreadsheet_files(): Returns a list of all spreadsheet files available to the authorized client.

# Worksheet Operations
# sheet.add_worksheet(title, rows, cols): Adds a new worksheet with the given title, number of rows, and columns.
# sheet.get_worksheet(index): Returns the worksheet at the given index (0-based).
# sheet.worksheet(title): Returns a worksheet by its title.
# sheet.worksheets(): Returns a list of all worksheets in the spreadsheet.
# sheet.del_worksheet(worksheet): Deletes the specified worksheet.
# sheet.duplicate_sheet(source_sheet_id): Duplicates the worksheet with the specified ID.

# Cell Operations
# sheet.acell(label): Returns the value of the cell with the specified label (e.g., A1).
# sheet.cell(row, col): Returns the value of the cell at the specified row and column.
# sheet.range(A1_range): Returns a list of cells in the specified range (e.g., A1:C10).
# sheet.get_all_values(): Returns all values from the worksheet as a list of lists.
# sheet.row_values(row): Returns all values in the specified row.
# sheet.col_values(col): Returns all values in the specified column.
# sheet.update_acell(label, value): Updates the value of the cell with the specified label.
# sheet.update_cell(row, col, value): Updates the value of the cell at the specified row and column.
# sheet.update(range, values): Updates the values in the specified range with a list of lists.

# Row and Column Operations
# sheet.append_row(values): Appends a row with the specified values to the end of the worksheet.
# sheet.insert_row(values, index): Inserts a row with the specified values at the given index.
# sheet.delete_rows(index): Deletes a row at the specified index.
# sheet.delete_columns(index): Deletes a column at the specified index.
# sheet.append_col(values): Appends a column with the specified values to the end of the worksheet.
# sheet.insert_col(values, index): Inserts a column with the specified values at the given index.

# Formatting and Batch Updates
# sheet.format(range, format): Applies formatting to the specified range. The format is provided as a dictionary (e.g., setting text color, background color).
# sheet.batch_update(body): Performs a batch update using the specified body, which is a dictionary containing the update requests.

# Data Finding and Sorting
# sheet.find(query): Finds the first cell matching the query.
# sheet.findall(query): Finds all cells matching the query.
# sheet.sort((column_index, direction)): Sorts the worksheet by the specified column index in the given direction (1 for ascending, 0 for descending).

# Permissions and Sharing
# sheet.share(email, perm_type, role): Shares the spreadsheet with the specified email address. perm_type can be 'user' or 'group', and role can be 'owner', 'writer', or 'reader'.
# sheet.list_permissions(): Returns a list of all permissions for the spreadsheet.
# sheet.remove_permissions(email): Removes permissions for the specified email address.