from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC


url_file_path = 'urls.txt'  # <--- IMPORTANT: You need to set this!

# Initialize the webdriver (e.g., Chrome)


try:
    # Read URLs from the text file
    with open(url_file_path, 'r') as file:
        urls = [line.strip() for line in file]  # Read each line, remove leading/trailing whitespace
    options = Options()
    driver = webdriver.Edge(options=options)
    # Iterate through the URLs
    for url in urls:
        print(f"Processing URL: {url}")
        driver.get(url)

        # Wait until a specific element is present on the page (adjust the timeout and element as needed)
        try:
            element = WebDriverWait(driver, 100).until(
                 EC.presence_of_element_located((By.CLASS_NAME, "home"))  # Example: wait for an element with ID "someElementId"
            )
            print(f"Page loaded successfully for {url}!")
        except Exception as wait_error:
            print(f"Page load failed for {url}: {wait_error}")


except FileNotFoundError:
    print(f"Error: URL file not found at {url_file_path}")
except Exception as e:
    print(f"An error occurred: {e}")

finally:
    # Close the browser
    driver.quit()