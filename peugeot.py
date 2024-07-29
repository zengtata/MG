import requests
from bs4 import BeautifulSoup
import json
from urllib.parse import urljoin

# Base URL of the page to scrape
base_url = 'https://www.peugeot.hu/valasszon-modellt/peugeot-kinalat.html'

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

# Find the specific <div> containing the data-resource-url
data_div = soup.find('div', class_='q-page-container grid-bg-transparent')
if not data_div:
    print('Could not find the main data container.')
    exit()

# Find the <div> with the data-resource-url attribute
results_div = soup.find('div', class_='q-mod q-mod-hmf-results q-hmf-results')
if not results_div:
    print('Could not find the results container.')
    exit()

# Extract the JSON URL from the data-resource-url attribute
json_url = results_div.get('data-resource-url')
if not json_url:
    print('No JSON URL found in data-resource-url attribute.')
    exit()

# Construct the full JSON URL
json_url = urljoin(base_url, json_url)

# Function to fetch and parse JSON data
def fetch_json_data(json_url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    response = requests.get(json_url, headers=headers)
    response.raise_for_status()
    return response.json()

# Fetch the JSON data
try:
    data = fetch_json_data(json_url)
    # Extract and print <li> elements from the JSON data
    try:
        # Assuming the JSON structure contains a 'data' field with the list items
        items = data.get('data', [])
        for item in items:
            # Print each item; you might need to adjust this based on the actual JSON structure
            print(json.dumps(item, indent=2))  # Print the whole item for clarity
    except KeyError as e:
        print(f"Error extracting data: {e}")
except Exception as e:
    print(f"Failed to fetch JSON data: {e}")

print('Data extraction complete')
