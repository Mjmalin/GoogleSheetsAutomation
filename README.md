# VeoSheetsAutomation

This program was made for the micromobility company Veo. It automates a time-consuming task in Google Sheets for their supply chain.

image

## The Problem

Each month, the Veo Lead Central Warehouse Manager has to print the parts list (and parts labels) every market location needs. There could be up to 70 parts lists and labels to print.

For each market, this entails:  
-manually updating the market location and central warehouse location in Google Sheets, waiting for the sheet to update  
-selecting the relevant section of the sheet to print  
-copy/pasting the info needed for parts labels into another sheet  
-selecting the relevant section of that sheet to print 

This program automates these tasks.

## How the Program Works

This program uses the gspread library to issue commands to the Google Sheets API.

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

The program finds the location of all cells that say "ok", as those are the markets that have shipments this month. 

```bash
# Find all cells in rows 3 and 4 of Supply Plan sheet that say "ok" or "OK"
cell_list = supply_plan_sheet.findall("ok")
cap_cell_list = supply_plan_sheet.findall("OK")
for caps in cap_cell_list:
    cell_list.append(caps)
```

The markets are added to one of two lists via the create_markets_lists function, depending on which central warehouse it will receive the shipment from. 

```bash
# For all matching cells, add corresponding market to a list for that location
def create_markets_lists(row, markets_list):
    for cell in cell_list:
        if cell.row == row:
            x = 2
            y = cell.col
            markets = supply_plan_sheet.cell(x, y).value
            markets_list.append(markets)
    markets_list.sort()
    print(markets_list)
```
The prepare_markets_pdfs function starts by updating the market and warehouse location in Sheets, which updates the spreadsheet with the parts and labels for that market for the month.

```bash
# Prepare the market PDFs for download
def prepare_market_pdfs(markets_list, location):
    for markets in markets_list:
        flat_list.clear()
        
        # Update market and location
        order_creator_sheet.update_cell(2, 14, location)
        order_creator_sheet.update_cell(1, 14, markets)
```
The function defines the number of rows to be printed as 'counts.'

```bash
        # Define proper rows range for each market, only as far as Special Requests that actually have contents
        values_list = order_creator_sheet.get('C:C')
        counts = 0
        for lines in values_list:
            counts = counts + 1
```
The function reads the column of parts labels and transfers them to a specially formatted sheet, as the labels need to be printed onto a certain size of physical sticky labels. To avoid Google Sheets API's 60 read/write requests per user per minute, the program is instructed to sleep when necessary.

```bash
        # Read labels (do not pull special request labels)
        non_sr_label_counts = 0
        counting_label_list = order_creator_sheet.get('T5:T')
        for non_sr_labels in counting_label_list:
            non_sr_label_counts = non_sr_label_counts + 1
        labels_list = order_creator_sheet.get(f'X5:X{non_sr_label_counts}')
        
        # Avoid google sheets API quota limit of 60 read/write requests per user per minute
        if len(labels_list) > 20:
            time.sleep(60)

        # Flatten the list of lists    
        for labels in labels_list:   
            for flat_labels in labels:
                flat_list.append(flat_labels)
        
        # Write labels in 'J' (10th) column
        labels_count = 0
        for flat_labels in flat_list:    
            labels_count = labels_count + 1
        if len(flat_list > 40):
                time.sleep(1)
        label_creator_sheet.update_cell(labels_count, 10, flat_labels)
```
Similar to the parts lists, the correct amount of rows to be printed also needs to be defined for the labels.

 ```bash       
        # Divide amount of labels by 3, rounded up, and pull that many rows of labels
        labels_values_list = label_creator_sheet.get('J:J')
        values_counts = 0.0
        for values in labels_values_list:
            values_counts = values_counts + 1.0
        div_values_counts = math.ceil(values_counts/3.0)
```
The function wraps up by feeding info into another function, which is used to download the parts and parts labels so they can be printed.

```bash
        # Download markets PDFs
        download_pdfs(order_creator_sheet.title, 'A1:J', counts, order_creator_sheet.id, location, markets)
        
        # Download labels PDFs
        download_pdfs(label_creator_sheet.title, 'A1:E', div_values_counts, label_creator_sheet.id, location, f'{markets}_labels')
        
        # Clear labels
        label_creator_sheet.batch_clear(['J:J'])
```

That download_pdfs function downloads the specified ranges of parts or parts labels, and creates an authorized session from Google Sheets API to then download. 

```bash
def download_pdfs(sheet_name, cell_range, counts, sheet_id, location, markets):
        # Construct the export URL with the specified range
        range_to_export = f"{sheet_name}!{cell_range}{counts}"
        export_url = f"https://docs.google.com/spreadsheets/d/{sheets_id}/export?format=pdf&gid={sheet_id}&range={range_to_export}"

        # Use authorized session to make the request
        session = requests.Session()
        session.headers.update({'Authorization': f'Bearer {creds.token}'})

        # Make the request to download the PDF
        response = session.get(export_url)

        # Save the content to a PDF file
        with open(f"{location}_{markets}.pdf", 'wb') as f:
            f.write(response.content) 
        pdfs.append(f"{location}_{markets}.pdf")
        print(f"PDF created: {location}_{markets}.pdf")
```

Finally, the PDFs are merged to be easily printed. The individual files are not deleted in case it's useful to reference those.

```bash
# Merge PDFs to one file to be physically printed
result = fitz.open()
for pdf in pdfs:
    with fitz.open(pdf) as mfile:
        result.insert_pdf(mfile)    
result.save("Markets.pdf")
print("PDF created: Markets.pdf")
```












