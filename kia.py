import requests
from bs4 import BeautifulSoup
import os
from urllib.parse import urljoin

# URL of the page to scrape
base_url = 'https://www.kia.com/hu/vasarlas/arlistak-katalogusok-adatok/'


# Function to download a file from a given URL
def download_file(url, output_folder):
    local_filename = os.path.join(output_folder, url.split('/')[-1])
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
    return local_filename


# Make a request to the website
response = requests.get(base_url)
response.raise_for_status()  # Check if the request was successful

# Parse the HTML content
soup = BeautifulSoup(response.text, 'html.parser')

# Find all <a> elements
links = soup.find_all('a')

# Create a folder to save the PDFs
output_folder = 'kia'
os.makedirs(output_folder, exist_ok=True)

# Loop through all <a> elements and check if they link to a PDF in the specified directory
for link in links:
    href = link.get('href')
    if href and href.endswith(
            '.pdf') and 'content/dam/kwcms/kme/hu/hu/assets/contents/utility/Brochure/price-list/' in href:
        # Construct the full URL if the href is relative
        pdf_url = urljoin(base_url, href)
        try:
            print(f'Downloading {pdf_url}')
            download_file(pdf_url, output_folder)
            print(f'Successfully downloaded {pdf_url}')
        except Exception as e:
            print(f'Failed to download {pdf_url}: {e}')

print('Download complete')
