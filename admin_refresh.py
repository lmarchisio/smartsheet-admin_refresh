import smartsheet
import logging
from playsound import playsound
import timeit

# Refreshes Admin columns in a specific sheet
# Takes sheet ID as an input

# Set API access token
access_token = None

column_map = {} # a place to map column names to column IDs

# a list of department names used to create updates for the IMS columns
departments = ['CNC',
               'Design',
               'Fab',
               'Install',
               'Metal',
               'Paint',
               'Sculpt',
               'Shipping']

# initialize client
ss_client = smartsheet.Smartsheet(access_token)
# make sure we don't miss any errors
ss_client.errors_as_exceptions(True)
# log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# choose sheet
print('This is the script that performs a refresh on the admin columns.')
print('What sheet shall we update? (Sheet ID)')
this_sheet = input()

# start timer
start = timeit.default_timer()

# get sheet
sheet = ss_client.Sheets.get_sheet(this_sheet)
print('Loaded ' + str(len(sheet.rows))+ ' rows from sheet: ' + sheet.name)

# Build column map
for column in sheet.columns:
    column_map[column.title] = column.id

# function to create row updates for Start IMS columns
def make_start(source_row, source_department):
    # build new cell value
    new_cell = ss_client.models.Cell()
    new_cell.column_id = column_map[department + ' Start']
    new_cell.formula = '=IF([Labor / Complete]@row = 0, IF([Dept.]@row = "' + department + '", Start@row))'
    new_cell.strict = False

    # build new row to update
    new_row = ss_client.models.Row()
    new_row.id = source_row.id
    new_row.cells.append(new_cell)

    return new_row

# function to create row updates for Finish IMS columns
def make_finish(source_row, source_department):
    # build new cell value
    new_cell = ss_client.models.Cell()
    new_cell.column_id = column_map[department + ' Finish']
    new_cell.formula = '=IF([Labor / Complete]@row = 0, IF([Dept.]@row = "' + department + '", Finish@row))'
    new_cell.strict = False

    # build new row to update
    new_row = ss_client.models.Row()
    new_row.id = source_row.id
    new_row.cells.append(new_cell)

    return new_row

# function to create row updates for Approved? column
def make_approved(source_row):
    # build new cell value
    new_cell = ss_client.models.Cell()
    new_cell.column_id = column_map['Approved?']
    new_cell.formula = '=IF(OR([Supervisor Confirmed]@row = 1, [PM Override]@row = 1), 1, 0)'
    new_cell.strict = False

    # build new row to update
    new_row = ss_client.models.Row()
    new_row.id = source_row.id
    new_row.cells.append(new_cell)

    return new_row

# for each department
# build then write Start column
# then build and write Finish column
for department in departments:
    rowsToUpdate = []
    for row in sheet.rows:
        rowToUpdate = make_start(row, department)
        rowsToUpdate.append(rowToUpdate)

    print('Writing ' + department + ' Start')
    updated_row = ss_client.Sheets.update_rows(
        this_sheet,
        rowsToUpdate)

    rowsToUpdate = []
    for row in sheet.rows:
        rowToUpdate = make_finish(row, department)
        rowsToUpdate.append(rowToUpdate)

    print('Writing ' + department + ' Finish')
    updated_row = ss_client.Sheets.update_rows(
        this_sheet,
        rowsToUpdate)

# build and write the Approved? column
rowsToUpdate = []
for row in sheet.rows:
    rowToUpdate = make_approved(row)
    rowsToUpdate.append(rowToUpdate)

print('Writing Approved?')
updated_row = ss_client.Sheets.update_rows(
    this_sheet,
    rowsToUpdate)

# end timer
stop = timeit.default_timer()

# Finish it
print('You made it this far.')
print('Good Job.')
print('Duration of Smartsheet stuff was: ', stop - start)
playsound('Gong.wav')



