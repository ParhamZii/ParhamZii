# I wrote this script to automate the process of downloading logs from APC NetBotz Rack Monitor 200 devices (I'm not sure about other models)
# I am using python 3.9.7 and selenium 4.0.0
# I am using the pyperclip library to copy the clipboard data and pandas to save it to an Excel file.
# I am using the datetime library to get the current date. 
# I am using the webdriver_manager library to automatically download the Chrome WebDriver.
# I am using the Chrome WebDriver to interact with the APC UPS web interface.
# for windows users you might need to install pre-requisites: you can use this command on cmd:
# pip install pyperclip pandas datetime webdriver_manager selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
import pyperclip 
import datetime

# defining the variables:
USERNAME = 'username'
PASSWORD = 'password'
LOGIN_URL = 'http://your_apc_IP_ADDRESS/logon.htm'
current_date_persian = datetime.datetime.now().strftime("%Y%m%d")
# to save file in C partition you need to run the script as administrator and uncomment the following line and comment the next line
#OUTPUT_FILE = r"C:\apc-{current_date_persian}.xlsx"
OUTPUT_FILE = f"apc-{current_date_persian}.xlsx"

# defining the Function:
def apc_logs():

    # Step 1: Set up Chrome WebDriver automatically
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get(LOGIN_URL)
    try:
        # Step 2: Log in
        print("Logging in...")
        driver.find_element(By.NAME, "login_username").send_keys(USERNAME)
        driver.find_element(By.NAME, "login_password").send_keys(PASSWORD)
        driver.find_element(By.NAME, "submit").click()

        # Step 3: Wait for login to complete
        time.sleep(3)
        print("Login successful!")

        # Step 4: Navigate to 'Logs' tab
        print("Navigating to Logs tab...")
        driver.find_element(By.LINK_TEXT, "Logs").click()
        time.sleep(3)

        # Step 5: Click on Data > Log
        print("Navigating to Data > Log...")
        driver.find_element(By.XPATH,'//*[@id="navcontainer"]/table/tbody/tr/td/ul/li[2]/ul/li[1]/a').click()
        print("Waiting for 1 minute to load logs...")
        time.sleep(15)

        # Step 6: Navigate to 'launcn log in new widow' tab
        print("Navigating to launch Logs tab...")
        driver.find_element(By.XPATH, '//*[@id="content"]/table/tbody/tr[2]/td[2]/form/table/tbody/tr[6]/td/input[3]').click()
        
        # Step 7: Navigate to tab
        print("Switching to new tab...")
        driver.switch_to.window(driver.window_handles[1])
        time.sleep(30)

        # Step 8: Target the second table and right-click on a cell
        actions = ActionChains(driver)
        print("Right-clicking on a table cell...")
        tables = driver.find_elements(By.TAG_NAME, "table")  # Get all tables
        data_table = tables[1]  # Select the second table
        cells = data_table.find_elements(By.TAG_NAME, "td")
        actions.context_click(cells[30]).perform()  # Right-click on a random data cell
        time.sleep(2)

        # Step 9: Simulate 'Ctrl + A' (Select All) and 'Ctrl + C' (Copy)
        print("Simulating 'Ctrl + A' and 'Ctrl + C'...")
        actions.key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
        time.sleep(1)
        actions.key_down(Keys.CONTROL).send_keys('c').key_up(Keys.CONTROL).perform()
        time.sleep(2)

        # Step 10: Retrieve clipboard contents
        print("Retrieving clipboard contents...")
        clipboard_data = pyperclip.paste()
        if not clipboard_data:
            raise Exception("Clipboard is empty. Copy operation failed.")

        # Step 11: Save clipboard data to Excel
        print("Saving data to Excel...")
        rows = [row.split("\t") for row in clipboard_data.split("\n") if row]
        df = pd.DataFrame(rows)
        df.to_excel(OUTPUT_FILE, index=False, header=False)
        print(f"Logs have been saved successfully to {OUTPUT_FILE}")

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # Step 12: Log out and close the browser
        try:
            print("Attempting to log out...")
            driver.switch_to.window(driver.window_handles[0])
            driver.find_element(By.LINK_TEXT, "Log out").click()
            print("Logged out successfully.")
        except Exception:
            print("Logout failed or was unavailable.")
        finally:
            driver.quit()
            print("Browser closed.")

if __name__ == "__main__":
    apc_logs()




       
  
