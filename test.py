from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Initialize the WebDriver (Chrome in this example)
driver = webdriver.Firefox()

# Navigate to the search page
driver.get('https://fcivietnam.com/du-an?q_keys=&q_loai=00000000-0000-0000-0000-000000000000&q_von=00000000-0000-0000-0000-000000000000&q_tinh=27&q_tinhtrang=00000000-0000-0000-0000-000000000000&q_year=0&size=100')

def save_page_as_html(url, filename):
    driver.get(url)
    time.sleep(2)  # Wait for the page to load
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(driver.page_source)

def process_search_results():
    # Loop through search result pages
    while True:
        # Find all the items in the current result page
        items = driver.find_elements(By.CSS_SELECTOR, '.fci-user')  # Adjust selector as needed

        # Iterate over each item and save the page as HTML
        for index, item in enumerate(items):
            item_url = item.get_attribute('href')
            save_page_as_html(item_url, f'Downloads\\test\\result_{index}.html')

        # Check if there is a next page and navigate to it
        try:
            next_button = driver.find_element(By.CSS_SELECTOR, '.next-page-selector')  # Adjust selector as needed
            next_button.click()
            time.sleep(2)  # Wait for the next page to load
        except:
            break

process_search_results()

driver.quit()
