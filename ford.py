import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import os

# URL of the page to scrape
base_url = 'https://www.ford.hu/erdeklodoknek/ismerje-meg/katalogusok'

# Function to download a file from a given URL
def download_file(url, output_folder):
    local_filename = os.path.join(output_folder, url.split('/')[-1])
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Referer': base_url
    }
    with requests.get(url, headers=headers, stream=True) as r:
        r.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
    return local_filename

# Function to fetch and parse a page
def fetch_page(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return BeautifulSoup(response.text, 'html.parser')

# Fetch the main page
print(f'Fetching the main page: {base_url}')
soup = fetch_page(base_url)
print('Main page fetched successfully')

# Find all <a> elements with href containing the specific pattern
brochure_links = soup.find_all('a', href=lambda href: href and '/content/overlays/download-a-brochure-3-0/' in href)
print(f'Found {len(brochure_links)} brochure links on the main page')

# Create a folder to save the PDFs
output_folder = 'ford'
os.makedirs(output_folder, exist_ok=True)

# Loop through all found <a> elements and process the second level links
for link in brochure_links:
    href = link.get('href')
    if href:
        # Construct the full URL if the href is relative
        brochure_url = urljoin(base_url, href)
        try:
            print(f'Processing brochure page: {brochure_url}')
            brochure_soup = fetch_page(brochure_url)
            # Find all <a> elements with href containing the PDF pattern
            pdf_links = brochure_soup.find_all('a', href=lambda href: href and '/content/dam/guxeu/hu/hu_hu/documents/pricelists/cars/' in href and href.endswith('.pdf'))
            print(f'Found {len(pdf_links)} PDF links on the brochure page: {brochure_url}')
            for pdf_link in pdf_links:
                pdf_href = pdf_link.get('href')
                if pdf_href:
                    pdf_url = urljoin(base_url, pdf_href)
                    try:
                        print(f'Downloading {pdf_url}')
                        download_file(pdf_url, output_folder)
                        print(f'Successfully downloaded {pdf_url}')
                    except Exception as e:
                        print(f'Failed to download {pdf_url}: {e}')
        except Exception as e:
            print(f'Failed to process {brochure_url}: {e}')

print('Download complete')
