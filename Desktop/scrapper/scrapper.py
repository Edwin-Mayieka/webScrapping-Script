from selenium import webdriver
import time
import datetime
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import openpyxl
import os
# Set the path to the Chrome binary
chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = '/usr/bin/chromium'  # Replace with the actual path to the Chrome binary

# Set up the Chrome driver with the downloaded ChromeDriver executable
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

#= Navigate to the Schoology login page
driver.get('https://app.schoology.com/')

# Find the username, password, and school code fields and fill them in
username_field = driver.find_element(By.ID, 'edit-mail')
password_field = driver.find_element(By.ID, 'edit-pass')
school_code_field = driver.find_element(By.ID, 'edit-school')
school_id_field=driver.find_element(By.ID, 'edit-school-nid')
school_code = '827110743'

username_field.send_keys('A00267623')
password_field.send_keys('090502')
time.sleep(3) 
password_field.click()
driver.execute_script("arguments[0].value = arguments[1];", school_id_field, school_code)
# Submit the login form
driver.find_element(By.ID, 'edit-submit').click()
driver.implicitly_wait(10)
driver.get('https://app.schoology.com/course/6667445896/members')
# driver.back()
# wait = WebDriverWait(driver, 10)

contact_info={}
emails = []
phone_number=[]
# Find the table rows (excluding the table header)
def loadData(clicks):
    clicks=clicks
    wait = WebDriverWait(driver, 10)
    table = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'table[role="presentation"]')))
    ids=[]
    rows = table.find_elements(By.TAG_NAME, 'tr')
    for row in rows:
        user_id = row.get_attribute('id')
        ids.append(user_id)     
    for id in ids:
        if id:
            email=''
            phone=''
            # Construct the user page URL using the extracted user ID
            user_page_url = f'https://app.schoology.com/user/{id}/info' 
            # Navigate to the user page
            driver.get(user_page_url)
            # Find the email element and extract the email address
            try:
                email_element = driver.find_element(By.CSS_SELECTOR, 'a.sExtlink-processed.mailto')
                email = email_element.get_attribute('href').replace('mailto:', '')
                emails.append(email)
                contact_info.update({email:'phone'})
            except NoSuchElementException:
                if email:
                    contact_info.update({email:''})
                    
                else:
                    continue
    driver.get('https://app.schoology.com/course/6667445896/members')
    wait = WebDriverWait(driver, 10)

    def is_page_stale(driver):
        try:
        # Check if the current page has become stale
            driver.refresh()
            return False
        except:
        # If any exception occurs, assume the page is stale
            return True

    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.next')))
    next_page = driver.find_element(By.CSS_SELECTOR, "div.next")
    clicks = clicks

    for click in range(clicks):
        next_page.click()
        print(f"clicked {click+1} times")

    # Wait for the page to become stale (indicating a new page has been loaded)
        wait.until(lambda driver: is_page_stale(driver))

    # Wait for the new "next page" element to be present before proceeding
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.next')))
        next_page = driver.find_element(By.CSS_SELECTOR, "div.next")

    
def print_dict_to_excel(dictionary, filename):
    # Create a new workbook and select the active sheet
    current_dir = os.getcwd()
    file_path = os.path.join(current_dir, filename)
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write the key-value pairs to the sheet
    row = 1
    for key, value in dictionary.items():
        sheet.cell(row=row, column=1).value = key
        sheet.cell(row=row, column=2).value = value
        row += 1

    # Save the workbook
    workbook.save(file_path)
    print("Data printed to Excel successfully.") 

def updatePage(pages):
    for i in range(pages):
        clicks=i+1
        loadData(clicks)
        print(i)         
# Print the list of email addresses
current_time = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
filename='contact_info_'+ current_time + '.xlsx'
# loadData()
# print(contact_info)

updatePage(3)
print_dict_to_excel(contact_info, filename=filename)

# Close the browser
driver.quit()
