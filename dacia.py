import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

# URL of the page to scrape
url = 'https://www.dacia.hu/arlista-letoltes.html'

# Make a request to the website
response = requests.get(url)
response.raise_for_status()  # Check if the request was successful

# Parse the HTML content
soup = BeautifulSoup(response.text, 'html.parser')

# Find all the specified <div> elements
divs = soup.find_all('div', class_='col-xs-12 col-sm-6 col-sm-6--clear-third col-md-4')

pdf_links = []

# Create the dacia directory if it doesn't exist
os.makedirs('dacia', exist_ok=True)

# Function to download a file
def download_file(url, folder):
    local_filename = os.path.join(folder, url.split('/')[-1])
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
    return local_filename

# Loop through each <div> and find the <a> element with href ending in 'price.pdf'
for div in divs:
    a_tag = div.find('a', href=True)
    if a_tag and 'price.pdf' in a_tag['href']:
        pdf_url = urljoin(url, a_tag['href'])
        pdf_links.append(pdf_url)

# Download the PDFs found
if pdf_links:
    for pdf_link in pdf_links:
        try:
            print(f"Downloading {pdf_link}")
            download_file(pdf_link, 'dacia')
            print(f"Successfully downloaded {pdf_link}")
        except Exception as e:
            print(f"Failed to download {pdf_link}: {e}")
else:
    print("No PDF links found ending with 'price.pdf'.")

print('Download complete.')
