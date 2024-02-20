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
base_url = "https://mosbatesabz.com/product-category/marriage/"

# Number of pages to scrape
num_pages = 150  # Change this to the desired number of pages

# Loop through pages
for page_num in range(1, num_pages + 1):
    try:
        # Construct the URL for the current page
        url = f"{base_url}page/{page_num}/"

        # Make a request to the website
        response = requests.get(url)
        response.raise_for_status()  # Raise an exception for unsuccessful HTTP requests

        soup = BeautifulSoup(response.text, 'html.parser')

        # Find and extract all product items
        product_items = soup.find_all('h3', class_='wd-entities-title')
        price_items = soup.find_all('span', class_='price')

        # Extract and write the text within each <h3> tag as the title
        # Extract and write the price within each <span> tag as the price
        for product_item, price_item in zip(product_items, price_items):
            try:
                product_title = product_item.find('a').text.strip() if product_item.find('a') else "Product Title Not Found"
                price = price_item.find('bdi').text.strip().replace('تومان', '').replace(',', '').strip() if price_item.find('bdi') else "Price Not Found"
                # Write data to the Excel worksheet
                # Ensure to handle the case where price is not available
                if price.isdigit():
                    worksheet.append([product_title, int(price.replace('&nbsp;', ''))])
                else:
                    worksheet.append([product_title, price])
            except Exception as e:
                print(f"Error scraping product details: {e}")

    except requests.exceptions.RequestException as e:
        print(f"Error fetching page {page_num}: {e}")

# Save the Excel file
workbook.save('SexualHygiene.xlsx')

print("Data has been successfully scraped and saved to 'SexualHygiene.xlsx'")
