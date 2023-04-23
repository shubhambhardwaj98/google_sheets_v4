# google_sheets_v4
Python Script to perform basic CRUD in Google Sheets using the Google Sheets V4 API


Generating credentials.json :
https://developers.google.com/sheets/api/quickstart/python

Inputs Required - 
  1. credentials.json - used to authenticate and generate the pickle file
  2. SPREADSHEET_ID = present in the URL of your google sheet 
    https://docs.google.com/spreadsheets/d/1s6k4T0mq2F4QdMWfJWX4JbYBQLg5x1nadpNTvnzOIMk/edit#gid=0
    
    SPREADSHEET_ID = 1s6k4T0mq2F4QdMWfJWX4JbYBQLg5x1nadpNTvnzOIMk
          

Current Functionalites - 

1. Creating a sheet
2. Reading from a Google sheet and storing it in a Pandas Dataframe based on the sheet name provided.
3. Updating Data in a sheet -  Existing Data is overwritten
4. Update Append - Appends the new data alongside existing data based on an outer join.
5. Deleting a sheet
6. Sorting a sheet based on multiple columns.
