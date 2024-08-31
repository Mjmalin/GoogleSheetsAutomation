import gspread
from google.oauth2.service_account import Credentials
import requests
import fitz
import math
import time

cwh_markets_list = list()
cwh2_markets_list = list()
pdfs = list()
flat_list = list()

# Authenticate and connect to Google Sheets
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

# Open the specific Google Sheet and worksheet
sheets_id = "1jWE3hhSSMi6wGoxvkd61owjmuZTpbzRXXupKNY4G39Y"
sheet = client.open_by_key(sheets_id)
order_creator_sheet = sheet.worksheet('Order Creator - Supply Plan')
supply_plan_sheet = sheet.worksheet('Supply Plan')
label_creator_sheet = sheet.worksheet('Label Creator - Pick / Pack')
        
# Find all cells in rows 3 and 4 of Supply Plan sheet that say "ok"
cell_list = supply_plan_sheet.findall("ok")

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
    
def download_pdfs(sheet_name, cell_range, counts, sheet_id, location, markets):
        # Construct the export URL with the specified range
        range_to_export = f"{sheet_name}!{cell_range}{counts}"
        export_url = f"https://docs.google.com/spreadsheets/d/{sheets_id}/export?format=pdf&gid={sheet_id}&range={range_to_export}"

        # Use authorized session to make the request (is this necessary?)
        session = requests.Session()
        session.headers.update({'Authorization': f'Bearer {creds.token}'})

        # Make the request to download the PDF
        response = session.get(export_url)

        # Save the content to a PDF file
        with open(f"{location}_{markets}.pdf", 'wb') as f:
            f.write(response.content) 
        pdfs.append(f"{location}_{markets}.pdf")
        print(f"PDF created: {location}_{markets}.pdf")

# Prepare the market PDFs for download
def prepare_market_pdfs(markets_list, location):
    for markets in markets_list:
        flat_list.clear()
        
        order_creator_sheet.update_cell(2, 14, location)
        order_creator_sheet.update_cell(1, 14, markets)

        # Define proper rows range to print for each market
        values_list = order_creator_sheet.get('C:C')
        counts = 0
        for lines in values_list:
            counts = counts + 1

        # Create labels
        labels_list = order_creator_sheet.get('S5:S')

        # Flatten the list of lists    
        for labels in labels_list:   
            for flat_labels in labels:
                flat_list.append(flat_labels)

        # Write labels
        labels_count = 0
        for flat_labels in flat_list:    
            if len(flat_list) > 20:
                time.sleep(1)
            labels_count = labels_count + 1
            label_creator_sheet.update_cell(labels_count, 10, flat_labels)
        
        # Define the amount of rows of labels to print
        div_values_counts = math.ceil(len(flat_list)/3.0)
            
        # Download market PDFs
        download_pdfs(order_creator_sheet.title, 'A1:J', counts, order_creator_sheet.id, location, markets)
        
        # Download labels PDFs
        download_pdfs(label_creator_sheet.title, 'A1:E', div_values_counts, label_creator_sheet.id, location, f'{markets}_labels')
        
        # Clear labels
        label_creator_sheet.batch_clear(['J:J'])
            
        
create_markets_lists(3, cwh_markets_list)

create_markets_lists(4, cwh2_markets_list)

print("Creating PDFs...")

prepare_market_pdfs(cwh_markets_list, "CWH")

prepare_market_pdfs(cwh2_markets_list, "CWH2")
    
# Merge PDFs to one file to be physically printed
result = fitz.open()
for pdf in pdfs:
    with fitz.open(pdf) as mfile:
        result.insert_pdf(mfile)    
result.save("Markets.pdf")
print("PDF created: Markets.pdf")
