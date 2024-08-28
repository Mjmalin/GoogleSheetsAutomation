# VeoSheetsAutomation

This program was made for the micromobility company Veo. It automates a time-consuming task in Google Sheets for their supply chain.

image

## The Problem

Each month, the Veo Central Warehouse Lead Manager has to print the parts list (and parts labels) every market location needs. There could be up to 70 parts lists and labels to print.

For each market, this entails:  
-manually updating the market location and central warehouse location in Google Sheets, waiting for the sheet to update  
-selecting the relevant section of the sheet to print  
-copy/pasting the info needed for parts labels into another sheet  
-selecting the relevant section of that sheet to print 

This program automates these tasks.

## How the Program Works

This program uses the library gspread to issue commands to the Google Sheets API.

```bash
import gspread
```

First, the program connects to the spreadsheet we need, using a key saved in a file "credentials.json" to authenticate. 

```bash
# Authenticate and connect to Google Sheets
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

# Open the Google Sheet and name the worksheets we need
sheets_id = "1jWE3hhSSMi6wGoxvkd61owjmuZTpbzRXXupKNY4G39Y"
sheet = client.open_by_key(sheets_id)
order_creator_sheet = sheet.worksheet('Order Creator - Supply Plan')
supply_plan_sheet = sheet.worksheet('Supply Plan')
label_creator_sheet = sheet.worksheet('Label Creator - Pick / Pack')
```









