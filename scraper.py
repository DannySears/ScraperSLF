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
from bs4 import BeautifulSoup

@dataclass
class Case:
    name: str
    address: str
    city: str
    state: str
    zip: str
    charge: str
    case_number: str
    attorney_present: bool  # True if attorney is present, False otherwise

def main():
    driver_path = os.path.join(os.path.dirname(__file__), 'drivers', 'chromedriver')

    while True:
        user_input = input("Desired access date in format mm/dd/yyyy: ")
        if validate_date(user_input):
            print("Valid date entered.")
            break
        else:
            print("Error: Date format should be mm/dd/yyyy. Please enter a new date.")
    
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

    #old date fetcher
    today = today - datetime.timedelta(days=3)
    d1 = today.strftime("%m/%d/%Y")

    datefiledAfterField = driver.find_element(By.ID, "DateFiledOnAfter")
    datefiledAfterField.send_keys(user_input)
    datefiledBeforeField = driver.find_element(By.ID, "DateFiledOnBefore")
    datefiledBeforeField.send_keys(user_input)

    #find and click search button
    searchButton = driver.find_element(By.ID, "SearchSubmit")
    searchButton.click()


    year = today.strftime("%Y")
    parsed_date = datetime.datetime.strptime(user_input, '%m/%d/%Y')
    year_input = parsed_date.year
    year_input = str(year_input)

    links = driver.find_elements(By.PARTIAL_LINK_TEXT, year_input)
    total_links = len(links)
    index = 0  # Start with the first link

    # store information about each client 
    cases = []

    # go through each page and 
    while index < total_links:
        # Re-find all elements to avoid stale element reference issues
        links = driver.find_elements(By.PARTIAL_LINK_TEXT, year)

        # Find case number
        caseNumb = links[index].text

        # Click the link at the current index
        links[index].click()

        # Wait for the page to load or for any other necessary action
        time.sleep(.2)
        
        try:
            # Check if the defendant's name is under "PIr11" or "PIr12"
            if driver.find_element(By.ID, "PIr01").text == "Defendant":
                nameD_element = driver.find_element(By.ID, "PIr11")
            else:
                nameD_element = driver.find_element(By.ID, "PIr12")

            nameD = nameD_element.text

            

            # Navigate to the address element
            # The address is in the 'td' element following the 'th' element for the defendant
            addressD_element = nameD_element.find_element(By.XPATH, "ancestor::tr/following-sibling::tr[1]/td")
            addressD = addressD_element.text

            td_html = addressD_element.get_attribute('innerHTML')

            soup = BeautifulSoup(td_html, 'html.parser')
            
            # Replace <br /> tags with newline characters
            for br in soup.find_all("br"):
                br.replace_with("\n")

            # Get the text and split by newlines
            sections = soup.get_text().split('\n')

            # Clean up special characters from each section
            cleaned_sections = [section.strip().encode('ascii', 'ignore').decode('ascii') for section in sections]

            addressRaw = addressD
            addressD = addressD.split("DL:")[0].strip()
            addressD = addressD.split("SID:")[0].strip()

            #find charge information
            charge_table = driver.find_element(By.XPATH,  "//th[contains(text(), 'Charges:')]")
            charge_element = charge_table.find_element(By.XPATH, "following::tr[1]/td[2]")
            chargeD = charge_element.text

            
            # Get the page source
            page_source = driver.page_source

            # Check if 'Pro Se' is in the page source
            attorney_presentD = "Pro Se" not in page_source

            parts = addressD.split()

            # Assuming the last two parts are state and zip code
            if len(parts) >= 2:
                state = parts[-2]
                zip_code = parts[-1]

                # The rest of the parts are the street address and city
                street_city = ' '.join(parts[:-2])
            else:
                # Fallback if the address doesn't have enough parts
                state = ''
                zip_code = ''
                street_city = addressD

            street_city = cleaned_sections[0]

            apt_indicators = {"APT", "UNIT", "SUITE", "BLDG", "#"}

            if contains_apt_indicator(addressRaw, apt_indicators):
                city = cleaned_sections[2].split(",")[0].strip()
                street_city = street_city + " " + cleaned_sections[1].split(",")[0].strip()
            else: 
                city = cleaned_sections[1].split(",")[0].strip()

            # Address tests
            print(len(cleaned_sections))
            print("Street Address:", street_city)
            print("City: ", city)
            print("State:", state)
            print("Zip Code:", zip_code)

            # Full output tests
            #print("Name:", nameD)
            #print("Address:", addressD)
            #print("Case number", caseNumb)
            #print("Charge", chargeD)
            #print("Attorney present", attorney_presentD)

        except Exception as e:
            print("An error occurred:", e)

        cases.append(Case(
            name=nameD,
            address=street_city,
            city=city,
            state=state,
            zip=zip_code,
            charge=chargeD,
            case_number=caseNumb,
            attorney_present=attorney_presentD
        ))  
        
        
        # If the link opens in the same tab and you need to go back to the original page
        driver.back()
        # Wait for the original page to load again
        time.sleep(.2)
        # Increment the index to click the next link in the next iteration
        index += 1

    # Convert the list of Case instances to a DataFrame
    df = pd.DataFrame([case.__dict__ for case in cases])

    formatted_date = user_input.replace("/", "-")

    # Export the DataFrame to an Excel file
    fileName = "CaseRecords" + formatted_date + ".xlsx"
    try:
        df.to_excel(fileName, index=False, engine='openpyxl')
    except Exception as e:
        print(f"An error occurred while saving the file: {e}")

    # Change columns width
    workbook = load_workbook(fileName)
    sheet = workbook.active
    sheet.column_dimensions['A'].width = 33
    sheet.column_dimensions['B'].width = 40
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 6
    sheet.column_dimensions['E'].width = 8
    sheet.column_dimensions['F'].width = 50
    sheet.column_dimensions['G'].width = 15
    sheet.column_dimensions['H'].width = 8
    workbook.save(fileName)

    # Output datapath to terminal
    excel_file_path = os.path.abspath(fileName)
    print(f"The Excel file is saved at: {excel_file_path}")

    time.sleep(10)
    driver.quit()

def contains_apt_indicator(address, apt_indicators):
    for indicator in apt_indicators:
        if indicator in address.upper():
            return True
    return False

def validate_date(date_string):
    try:
        datetime.datetime.strptime(date_string, '%m/%d/%Y')
        return True
    except ValueError:
        return False


if __name__ == '__main__':
    main()