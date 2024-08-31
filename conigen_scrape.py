# Instructions to operate script using VSCode:
# Please hit the run button at the top right corner of the window to operate the script (right facing triangle)
# A Chrome browser window will pop up, do not panic! This means the script is working, please leave the window open as it will close on its own
# Observe terminal for printed results after the script finishes running
# Navigate to your desktop to find a folder named: Acro_Product_List
# The relevant files should now be in the folder with the date in the name (YYYY-MM-DD)
# There will be two pages, Product List and New Products
import os
from datetime import date
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# chrome webdriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# open acro page
driver.get('https://www.acrobiosystems.com/A1337-GPCRs.html')

# wait for table to load
time.sleep(4)

# find the table rows under specific tbody tags
vlp_rows = driver.find_elements(By.CSS_SELECTOR, '#auto_vlp tr')
detergent_micelle_rows = driver.find_elements(By.CSS_SELECTOR, '#auto_detergent_micelle tr')
nanodisc_rows = driver.find_elements(By.CSS_SELECTOR, '#auto_nanodisc tr')

# initialize  list to hold the data
data = []

def extract_data(rows):
    for row in rows:
        # scrape cells within the row
        cells = row.find_elements(By.TAG_NAME, 'td')
        if len(cells) >= 3:
            molecule = cells[0].text.strip()
            product_number = cells[1].text.strip()
            product_name = cells[2].text.strip()

            # append the data as a tuple to the list
            data.append((molecule, product_number, product_name))

# extract data from each platform
extract_data(vlp_rows)
extract_data(detergent_micelle_rows)
extract_data(nanodisc_rows)

driver.quit()

# create a data frame
df_current = pd.DataFrame(data, columns=["Molecule", "Product Number", "Product Name"])


total_products = len(df_current)

desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

# defining path to send xlsx files
folder_name = "Acro_Product_List"
folder_path = os.path.join(desktop_path, folder_name)
os.makedirs(folder_path, exist_ok=True)

# setting date for file name and future comparison
today = date.today()
current_date = today.strftime("%Y-%m-%d")   # YYYY-MM-DD

# define  full file path with date in the file name
file_name = f"Acro_Products_{current_date}.xlsx"
file_path = os.path.join(folder_path, file_name)

# reverse sort previous files to find the most recent
previous_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and f.startswith('Acro_Products_')]
previous_files.sort(reverse=True)

if previous_files:
    latest_file = previous_files[0]
    latest_file_path = os.path.join(folder_path, latest_file)

    df_previous = pd.read_excel(latest_file_path, sheet_name='Product List')

    # find new products by comparing against old sheets using set operations
    current_set = set(map(tuple, df_current.values))
    previous_set = set(map(tuple, df_previous.values))
    new_products_set = current_set - previous_set
    new_products_df = pd.DataFrame(list(new_products_set), columns=df_current.columns)
else:
    print("No previous files found, proceeding with the current dataset.")
    new_products_df = df_current  # no comparison possible, so all products are considered new

# calculate statistics
total_products = len(df_current)
total_new_products = len(new_products_df)
percentage_new_products = (total_new_products / total_products) * 100 if total_products > 0 else 0

# print the results, can comment out if unecessary
print(f"Total number of products: {total_products}")
print(f"Total number of new products: {total_new_products}")
print(f"Percentage of new products: {percentage_new_products:.2f}%")

if total_new_products > 0:
    print("New Products:")
    print(new_products_df.to_string(index=False))

# prepare the Excel writer
with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    # write the new products data to the "New Products" sheet
    new_products_df.to_excel(writer, index=False, startrow=1, sheet_name='New Products')

    # write the full product list to the "Product List" sheet
    df_current.to_excel(writer, index=False, startrow=1, sheet_name='Product List')

    # Access the workbook and both worksheets
    workbook = writer.book
    worksheet_full = writer.sheets['Product List']
    worksheet_new = writer.sheets['New Products']

    # Apply a header format for the main header row
    main_header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'middle',
        'align': 'center',
        'border': 1,
        'bg_color': '#D7E4BC'
    })

    # format main header row
    for col_num, value in enumerate(df_current.columns.values):
        worksheet_full.write(1, col_num, value, main_header_format)
        worksheet_new.write(1, col_num, value, main_header_format)

    # set columns
    worksheet_full.set_column('A:A', 30)
    worksheet_full.set_column('B:B', 20)
    worksheet_full.set_column('C:C', 50)

    worksheet_new.set_column('A:A', 30)
    worksheet_new.set_column('B:B', 20)
    worksheet_new.set_column('C:C', 50)

    # statistics for percentage of new products
    total_row_full = len(df_current) + 2
    worksheet_full.write(total_row_full, 0, 'Total number of products:', workbook.add_format({'bold': True}))
    worksheet_full.write(total_row_full, 1, total_products)

    worksheet_full.write(total_row_full + 1, 0, 'Total number of new products:', workbook.add_format({'bold': True}))
    worksheet_full.write(total_row_full + 1, 1, total_new_products)

    worksheet_full.write(total_row_full + 2, 0, 'Percentage of new products:', workbook.add_format({'bold': True}))
    worksheet_full.write(total_row_full + 2, 1, f"{percentage_new_products:.2f}%")

    #stats
    total_row_new = len(new_products_df) + 2
    worksheet_new.write(total_row_new, 0, 'Total number of products:', workbook.add_format({'bold': True}))
    worksheet_new.write(total_row_new, 1, total_products)

    worksheet_new.write(total_row_new + 1, 0, 'Total number of new products:', workbook.add_format({'bold': True}))
    worksheet_new.write(total_row_new + 1, 1, total_new_products)

    worksheet_new.write(total_row_new + 2, 0, 'Percentage of new products:', workbook.add_format({'bold': True}))
    worksheet_new.write(total_row_new + 2, 1, f"{percentage_new_products:.2f}%")
print(f"Excel file saved successfully at: {file_path}")


