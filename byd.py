import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import os

# URL of the page to scrape
url = 'https://byd-wallismotor.hu/arlista/'

# Create the byd directory if it doesn't exist
os.makedirs('byd', exist_ok=True)

def download_pdf(pdf_url, folder):
    try:
        response = requests.get(pdf_url)
        response.raise_for_status()  # Check if the request was successful

        # Extract filename from the URL
        filename = os.path.basename(pdf_url)
        filepath = os.path.join(folder, filename)

        # Write the PDF content to a file
        with open(filepath, 'wb') as file:
            file.write(response.content)
        print(f"Downloaded: {filepath}")
    except Exception as e:
        print(f"Failed to download {pdf_url}: {e}")

# Make a request to the website
response = requests.get(url)
response.raise_for_status()  # Check if the request was successful

# Parse the HTML content
soup = BeautifulSoup(response.text, 'html.parser')

# Find the section with class "pricelist"
pricelist_section = soup.find('section', class_='pricelist')

if pricelist_section:
    # Find all <a> elements in the pricelist section
    links = pricelist_section.find_all('a', href=True)

    for link in links:
        href = link['href']
        # Check if the href contains '/uploads' and ends with '.pdf'
        if '/uploads' in href and href.endswith('.pdf'):
            # Construct the full URL of the PDF
            pdf_url = urljoin(url, href)
            print(f"PDF found: {pdf_url}")
            # Download the PDF
            download_pdf(pdf_url, 'byd')
else:
    print('No pricelist section found.')

print('Scraping complete.')
