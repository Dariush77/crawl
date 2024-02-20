import requests
from bs4 import BeautifulSoup
import openpyxl

# Create a new Excel workbook and add a worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active

# Write headers to the worksheet
worksheet['A1'] = 'Product Title'
worksheet['B1'] = 'Price (Toman)'

# Base URL for the product pages
base_url = "https://mosbatesabz.com/product-category/medical-equipment/page/"

# Make a request to the website to get the last page number
response = requests.get(base_url)
soup = BeautifulSoup(response.text, 'html.parser')

# Find the pagination section
pagination = soup.find('nav', class_='woocommerce-pagination')

# Find the last page link
last_page_link = pagination.find_all('a')[-2]  # Second to last link

# Extract the page number
last_page_number = int(last_page_link.text)

print("Last page number:", last_page_number)

# Loop through pages
for page_num in range(1, last_page_number + 1):
    # Construct the URL for the current page
    url = f"{base_url}{page_num}/"

    # Make a request to the website
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    # Find and extract all product names and prices
    product_items = soup.find_all('h3', class_='wd-entities-title')
    price_items = soup.find_all('span', class_='price')

    # Extract and write the text within each <h3> tag as the title
    # Extract and write the price within each <span> tag as the price
    for product_item, price_item in zip(product_items, price_items):
        product_title = product_item.find('a').text.strip()
        ins_tag = price_item.find('ins')
        if ins_tag:
            price = ins_tag.text.strip().replace('تومان', '').replace(',', '').strip()
        else:
            # If <ins> tag does not exist, try getting the price from <del> tag
            del_tag = price_item.find('del')
            if del_tag:
                price = del_tag.text.strip().replace('تومان', '').replace(',', '').strip()
            else:
                # If neither <ins> nor <del> tags exist, set price to None or handle the situation accordingly
                price = 'Price not available'
                
        # Write data to the Excel worksheet
        if price.isdigit():  # Check if price is a valid integer
            worksheet.append([product_title, int(price)])
        else:
            worksheet.append([product_title, price])

# Save the Excel file
workbook.save('Medicalequipment.xlsx')

print("Data has been successfully scraped and saved to 'Medicalequipment.xlsx'")
