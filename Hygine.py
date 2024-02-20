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
base_url = "https://mosbatesabz.com/product-category/beauty-and-personal-care/page/"

# Number of pages to scrape
num_pages = 150  # Change this to the desired number of pages

# Loop through pages
for page_num in range(1, num_pages + 1):
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
        price = price_item.find('bdi').text.strip().replace('تومان', '').replace(',', '').strip()

        # Write data to the Excel worksheet
        worksheet.append([product_title, int(price)])

# Save the Excel file
workbook.save('Hygine.xlsx')

print("Data has been successfully scraped and saved to 'Hygine.xlsx'")