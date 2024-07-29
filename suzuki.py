import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import os

# Base URL of the main page
main_url = 'https://auto.suzuki.hu/modellek'

def clean_text(text):
    """Remove non-breaking spaces and other unwanted characters from the text."""
    return text.replace('\xa0', ' ').strip()

# Function to extract model details and prices
def extract_model_data(model_url):
    try:
        # Fetch the model-specific page
        response = requests.get(model_url)
        response.raise_for_status()  # Check if the request was successful
        soup = BeautifulSoup(response.text, 'html.parser')

        # Extract price list table
        price_list_div = soup.find('div', class_='price-list-table-container')
        table = price_list_div.find('table')

        headers = []
        data = []

        if table:
            # Extract headers
            headers = [clean_text(th.text) for th in table.find_all('tr')[0].find_all('td')]
            # Extract rows
            rows = table.find_all('tr')[1:]  # Skip header row

            for row in rows:
                cols = [clean_text(col.text) for col in row.find_all('td')]
                if len(cols) > 1:  # Ensure row has data
                    data.append(cols)

        return headers, data

    except Exception as e:
        print(f"Failed to extract data from {model_url}: {e}")
        return [], []

# Fetch the main page
response = requests.get(main_url)
response.raise_for_status()  # Check if the request was successful
soup = BeautifulSoup(response.text, 'html.parser')

# Find all model links
model_links = soup.find_all('a', href=True)
model_urls = [urljoin(main_url, link['href']) for link in model_links if 'arlista' in link['href']]

# Create the suzuki directory if it doesn't exist
os.makedirs('suzuki', exist_ok=True)

# Prepare to write results to a text file
txt_filename = 'suzuki/suzuki_price_list.txt'

# Set to keep track of processed URLs
processed_urls = set()

# Write to text file
with open(txt_filename, mode='w', encoding='utf-8') as file:
    # Iterate through each model URL and extract data
    for model_url in model_urls:
        if model_url in processed_urls:
            print(f"Skipping already processed URL: {model_url}")
            continue

        try:
            print(f"Processing {model_url}")
            headers, rows = extract_model_data(model_url)
            if headers and rows:  # Only write to file if data extraction was successful
                # Write headers
                file.write(f"{' | '.join(headers)}\n")
                file.write('-' * 80 + '\n')  # Separator line

                # Write data rows
                for row in rows:
                    file.write(f"{' | '.join(row)}\n")

                file.write('\n' + '=' * 80 + '\n')  # End of current model section
                processed_urls.add(model_url)  # Mark this URL as processed
                print(f"Data from {model_url} written to text file")
        except Exception as e:
            print(f"Failed to process {model_url}: {e}")

print(f'Data extraction complete. Results saved to {txt_filename}')
