import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import os

# URL of the page to scrape
base_url = 'https://www.opel.hu/tools/arlistak-es-katalogusok.html'

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

# Function to fetch and parse the main page
def fetch_main_page(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return BeautifulSoup(response.text, 'html.parser')

# Fetch the main page
soup = fetch_main_page(base_url)

# Find all <a> elements with href containing the specific pattern
pdf_links = soup.find_all('a', href=lambda href: href and '/content/dam/opel/hungary/brochures/Pricelists/' in href and href.endswith('.pdf'))

# Create a folder to save the PDFs
output_folder = 'opel'
os.makedirs(output_folder, exist_ok=True)

# Loop through all found <a> elements and download the PDFs
for link in pdf_links:
    href = link.get('href')
    if href:
        # Construct the full URL if the href is relative
        pdf_url = urljoin(base_url, href)
        try:
            print(f'Downloading {pdf_url}')
            download_file(pdf_url, output_folder)
            print(f'Successfully downloaded {pdf_url}')
        except Exception as e:
            print(f'Failed to download {pdf_url}: {e}')

print('Download complete')
