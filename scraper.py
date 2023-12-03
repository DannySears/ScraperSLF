import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import datetime

def main():
    driver_path = os.path.join(os.path.dirname(__file__), 'drivers', 'chromedriver')

    #boot driver to home page
    driver = webdriver.Chrome()
    driver.get('https://cijspub.co.collin.tx.us/PublicAccess/CaseDetail.aspx?CaseID=2336564')

    #get the button for case records and click
    time.sleep(1)
    caseRecordsLink = driver.find_element(By.XPATH, "//div[@id='divOption1']/a[contains(@class, 'ssSearchHyperlink')]")
    caseRecordsLink.click()

    time.sleep(1)

    #click date filed button
    dateFiledButton = driver.find_element(By.ID, "DateFiled")
    dateFiledButton.click()
    time.sleep(1)

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



    time.sleep(10)
    driver.quit()

if __name__ == '__main__':
    main()