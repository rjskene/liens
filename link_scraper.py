import time
import os
import win32clipboard
from tqdm.auto import tqdm
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, StaleElementReferenceException

PROJECT_LINKS_FILE = 'static/project_links.csv'

class KoreTraxScraper:
    """
    A class to scrape data from the Koretrax website using Selenium and Chrome WebDriver.
    """
    
    def __init__(self, headless=False):
        """
        Initialize the scraper with Chrome WebDriver.
        
        Args:
            headless (bool): Whether to run Chrome in headless mode
        """
        self.main_url = "https://hts-texas.koretrax.com"
        self.driver = None
        self.headless = headless
        
    def setup_driver(self):
        """Set up and configure the Chrome WebDriver."""
        chrome_options = Options()
        if self.headless:
            chrome_options.add_argument("--headless")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.maximize_window()
            return True
        except WebDriverException as e:
            print(f"Error setting up Chrome WebDriver: {e}")
            return False
                
    def close(self):
        """Close the WebDriver and release resources."""
        if self.driver:
            self.driver.quit()
            self.driver = None

def navigate_to_project_page(scraper, project_id):
    """Navigate to the project page for a given project ID."""
    """
    Navigate to the project page for a given project ID by using the search functionality.
    
    Args:
        scraper (KoretraxScraper): The initialized scraper instance
        project_id (str): The project ID to search for
        
    Returns:
        bool: True if navigation successful, False otherwise
    """
    # Wait for the search input to be available
    search_input = WebDriverWait(scraper.driver, 10).until(
        EC.presence_of_element_located((By.ID, "launcherSearchInput"))
    )
    
    # Clear any existing text and enter the project ID
    search_input.clear()
    search_input.send_keys(project_id)

    search_input = WebDriverWait(scraper.driver, 10).until(
        EC.presence_of_element_located((By.ID, "launcherSearchInput"))
    ) # in case element goes stale
    # Press Enter to submit the search
    search_input.send_keys(Keys.RETURN)
    
    return True
    
def copy_project_link(scraper):
    """
    Copy the project link to the clipboard by clicking the domain more button
    and extracting the URL.
    
    Args:
        scraper (KoretraxScraper): The initialized scraper instance
        
    Returns:
        str: The project URL if successful, None otherwise
    """
    from tqdm.auto import tqdm
    
    progress = tqdm(total=4, desc="Starting link copy process", leave=False)
    
    # Find and click the domain more button
    more_button = WebDriverWait(scraper.driver, 10, ignored_exceptions=[StaleElementReferenceException]) \
        .until(EC.element_to_be_clickable((By.ID, "domainMoreButton")))
    more_button.click()
    progress.set_description("Clicked more button")
    progress.update(1)
    
    # Wait for the dropdown menu to appear and handle stale element exceptions
    dropdown_menu = WebDriverWait(scraper.driver, 10, ignored_exceptions=[StaleElementReferenceException]) \
        .until(EC.presence_of_element_located((By.ID, "domainMoreMenu")))
    
    
    # Find and click the "Show Link URL" option
    show_link_option = WebDriverWait(dropdown_menu, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Show Link URL')]"))
    )
    show_link_option.click()

    time.sleep(1)
    progress.set_description("Clicked show link url")
    progress.update(1)

    # Find and click the "Copy Link" button
    copy_link_button = WebDriverWait(scraper.driver, 10).until(
        EC.element_to_be_clickable((By.ID, "copyLinkButton"))
    )
    copy_link_button.click()
    progress.set_description("Clicked copy link button")
    progress.update(1)
    
    # Wait a moment for the clipboard operation to complete
    time.sleep(.5)
    
    # Open the clipboard and get the content
    win32clipboard.OpenClipboard()
    url = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    
    progress.set_description(f"Successfully copied URL: {url}")
    progress.update(1)
    progress.close()
    return url

def get_project_url(scraper, project_id):
    time.sleep(.5)
    navigate_to_project_page(scraper, project_id)
    url = copy_project_link(scraper)
    return url

def scrape_for_new_urls(needs_new_urls, existing_urls):
    # Validate existing_urls DataFrame structure
    if not all(col in existing_urls.columns for col in ["Project Number", "URL"]):
        raise ValueError("existing_urls DataFrame must contain only 'Project Number' and 'URL' columns")
    
    # Ensure existing_urls only has the required columns
    existing_urls = existing_urls[["Project Number", "URL"]]
    if needs_new_urls.shape[0] > 0:
        scraper = KoreTraxScraper()
        urls = []
        if scraper.setup_driver():
            try:
                scraper.driver.get(scraper.main_url)
                WebDriverWait(scraper.driver, 20).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
                progress = tqdm(total=len(needs_new_urls['Project Number']), desc="Getting project URLs")
                for project_id in needs_new_urls['Project Number']:
                    url = get_project_url(scraper, project_id)
                    urls.append((project_id, url))
                    progress.update(1)
            except Exception as e:
                import traceback
                print(f"Error occurred while scraping URLs:")
                print(traceback.format_exc())
                raise e from None
                
            finally:
                scraper.close()
                new_urls = pd.DataFrame(urls, columns=['Project Number', 'URL'])
                pd.concat([existing_urls, new_urls]).to_csv(PROJECT_LINKS_FILE, index=False)
                print (f"Successfully scraped {len(urls)} new URLs and saved to {PROJECT_LINKS_FILE}")
    else:
        print ('No new URLs needed')

 
def test_connection():
    """Test the connection to Koretrax website."""
    scraper = KoreTraxScraper()
    print("Setting up WebDriver...")
    if scraper.setup_driver():
        try:
            scraper.driver.get("https://hts-texas.koretrax.com")
            print("Successfully connected to Koretrax website!")
            return True
        except Exception as e:
            print(f"Failed to connect to Koretrax website: {e}")
            return False
        finally:
            scraper.close()
    else:
        print("Failed to set up WebDriver")
        return False
    
def test_navigate_to_project_page():
    """Test the navigation to the project page."""
    scraper = KoreTraxScraper()
    if scraper.setup_driver():
        try:
            scraper.driver.get(scraper.main_url)
            time.sleep(3)
            navigate_to_project_page(scraper, "21600013-DXS-1")
            copy_project_link(scraper)
        finally:
            scraper.close()
    else:
        print("Failed to set up WebDriver")

def check_or_create_project_csv():
    """Check or create the project CSV file."""
    # Check if the file exists
    if os.path.exists("project_links.csv"):
        print("Project CSV file already exists.")
        return
    else:
        print("Project CSV file does not exist. Creating new file.")
        with open("project_links.csv", "w") as f:
            f.write("project_id,project_link\n")
        return

if __name__ == "__main__":
    # test_connection()
    # test_navigate_to_project_page()
    check_or_create_project_csv()
