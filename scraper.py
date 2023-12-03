import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

driver_path = os.path.join(os.path.dirname(__file__), 'drivers', 'chromedriver')

driver = webdriver.Chrome()
driver.get('https://cijspub.co.collin.tx.us/PublicAccess/CaseDetail.aspx?CaseID=2336564')


tables = driver.find_elements(By.TAG_NAME, 'table')

defendant_name = driver.find_element(By.XPATH, "//table")
print(defendant_name)
# Check if there are at least three tables
if len(tables) >= 3:
    # Select the third table
    third_table = tables[4]

    # Find the first 'tr' element in the third table
    # If you need a specific 'tr', you can adjust this part
    tr_element = third_table.find_elements(By.TAG_NAME,'tr')

    rows = third_table.find_elements(By.TAG_NAME, 'tr')
    
    for row in rows:
    # Find all 'td' elements in the row
        cells = row.find_elements(By.TAG_NAME, 'td')
    for cell in cells:
        print(cell.text)

# Don't forget to close the driver
driver.quit()