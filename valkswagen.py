from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Configure the WebDriver
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run in headless mode (no GUI)
service = Service('path/to/chromedriver')  # Replace with the path to your chromedriver executable
driver = webdriver.Chrome(service=service, options=chrome_options)

# Open the webpage
url = 'https://konfigurator.volkswagen.hu/cc-hu/hu_HU_VW22/V/models'
driver.get(url)

# Extract and print all <span> elements
spans = driver.find_elements(By.TAG_NAME, 'span')
if spans:
    print("Spans found on the page:")
    for span in spans:
        print(span.text)
else:
    print("No <span> elements found.")

# Extract and print all <a> elements
links = driver.find_elements(By.TAG_NAME, 'a')
if links:
    print("\nLinks found on the page:")
    for link in links:
        print(link.get_attribute('href'))
else:
    print("No <a> elements found.")

# Close the WebDriver
driver.quit()
