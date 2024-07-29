import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

# URL of the page to scrape
url = 'https://www.toyota.hu/ajanlatok/arlistak'

# Make a request to the website
response = requests.get(url)
response.raise_for_status()  # Check if the request was successful

# Parse the HTML content
soup = BeautifulSoup(response.text, 'html.parser')

# Use CSS selectors to find the relevant elements
# Update selector based on actual structure
pdf_links = soup.select('div.models a[href*="/arlista_"][href$=".pdf"]')

if pdf_links:
    for link in pdf_links:
        # Get the href attribute
        href = link['href']
        # Construct the full URL if href is relative
        pdf_url = urljoin(url, href)
        # Print the name of the PDF
        pdf_name = href.split('/')[-1]
        print(f'PDF found: {pdf_name} - URL: {pdf_url}')
else:
    print('No PDFs found.')

print('Scraping complete.')
