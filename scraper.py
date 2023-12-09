import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import datetime
import pandas as pd
from dataclasses import dataclass
from openpyxl import load_workbook

@dataclass
class Case:
    name: str
    address: str
    charge: str
    case_number: str
    attorney_present: bool  # True if attorney is present, False otherwise

def main():
    driver_path = os.path.join(os.path.dirname(__file__), 'drivers', 'chromedriver')

    #boot driver to home page
    driver = webdriver.Chrome()
    driver.get('https://cijspub.co.collin.tx.us/SecurePA/Login.aspx?ReturnUrl=%2FSecurePA%2Fdefault.aspx')

    #get the button for case records and click
    time.sleep(.5)

    userName = 'ssears'
    password = 'ssears'

    #enter user and password
    userNameField = driver.find_element(By.ID, "UserName")
    userNameField.send_keys(userName)
    passwordField = driver.find_element(By.ID, "Password")
    passwordField.send_keys(password)
    signOnButton = driver.find_element(By.NAME, 'SignOn')
    signOnButton.click()
    
    #click case records
    caseRecordsLink = driver.find_element(By.XPATH, "//div[@id='divOption1']/a[contains(@class, 'ssSearchHyperlink')]")
    caseRecordsLink.click()

    time.sleep(.5)

    #click date filed button
    dateFiledButton = driver.find_element(By.ID, "DateFiled")
    dateFiledButton.click()
    time.sleep(.5)

    #enter the current date to the date filed fields
    today = datetime.date.today()

    today = today - datetime.timedelta(days=1)
    d1 = today.strftime("%m/%d/%Y")

    datefiledAfterField = driver.find_element(By.ID, "DateFiledOnAfter")
    datefiledAfterField.send_keys(d1)
    datefiledBeforeField = driver.find_element(By.ID, "DateFiledOnBefore")
    datefiledBeforeField.send_keys(d1)

    #find and click search button
    searchButton = driver.find_element(By.ID, "SearchSubmit")
    searchButton.click()


    year = today.strftime("%Y")

    links = driver.find_elements(By.PARTIAL_LINK_TEXT, year)
    total_links = len(links)
    index = 0  # Start with the first link

    # store information about each client 
    cases = []

    # go through each page and 
    while index < total_links:
        # Re-find all elements to avoid stale element reference issues
        links = driver.find_elements(By.PARTIAL_LINK_TEXT, "2023")

        # Click the link at the current index
        links[index].click()

        # Wait for the page to load or for any other necessary action
        time.sleep(.2)
        
        if driver.find_element(By.ID, "PIr01").text == "Defendant":
            nameD = driver.find_element(By.ID, "PIr11").text
        else: 
            nameD = driver.find_element(By.ID, "PIr12").text

        cases.append(Case(
            name=nameD,
            address="123 Main St",
            charge="Charge Details",
            case_number="001-XXXX-2023",
            attorney_present=True
        ))  
        
        
        # If the link opens in the same tab and you need to go back to the original page
        driver.back()
        # Wait for the original page to load again
        time.sleep(.2)
        # Increment the index to click the next link in the next iteration
        index += 1

    # Convert the list of Case instances to a DataFrame
    df = pd.DataFrame([case.__dict__ for case in cases])

    # Export the DataFrame to an Excel file
    fileName = "CaseRecords" + today.strftime("%m-%d-%Y") + ".xlsx"
    df.to_excel(fileName, index=False, engine='openpyxl')  

    # Change columns width
    workbook = load_workbook(fileName)
    sheet = workbook.active
    sheet.column_dimensions['A'].width = 33
    sheet.column_dimensions['B'].width = 40
    sheet.column_dimensions['C'].width = 56
    sheet.column_dimensions['D'].width = 25
    sheet.column_dimensions['E'].width = 10
    workbook.save(fileName)

    # Output datapath to terminal
    excel_file_path = os.path.abspath(fileName)
    print(f"The Excel file is saved at: {excel_file_path}")

    time.sleep(10)
    driver.quit()

if __name__ == '__main__':
    main()