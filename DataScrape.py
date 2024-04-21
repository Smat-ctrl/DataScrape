import time
import openpyxl
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException

options = Options()
options.add_experimental_option("detach", True)  # Keeps the window open

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

driver.get(
    "https://www.wcaworld.com/Directory?siteID=24&au=m&pageIndex=1&pageSize=100&searchby=CountryCode&country=DZ&city=&keyword=&orderby=CountryCity&networkIds=1&networkIds=2&networkIds=3&networkIds=4&networkIds=61&networkIds=98&networkIds=108&networkIds=118&networkIds=5&networkIds=22&networkIds=13&networkIds=18&networkIds=15&networkIds=16&networkIds=38&networkIds=103&layout=v1&submitted=search")
driver.maximize_window()

# Wait for the page to load before locating elements
driver.implicitly_wait(50) 

# Click the link that opens the login form
x_path_expression = '//*[@id="navright"]/div[2]/a'
link = driver.find_element(By.XPATH, x_path_expression)
link.click()

# Fill in the username and password inputs
username_input = driver.find_element(By.XPATH, '//*[@id="usr"]')
username_input.send_keys("") # add your username for the site

password_input = driver.find_element(By.XPATH, '//*[@id="pwd"]')
password_input.send_keys("") # add your own password

# Click the submit button (login button)
submit = driver.find_element(By.XPATH, '//*[@id="login-form-button"]')
submit.click()

directory = driver.find_element(By.XPATH, '//*[@id="navbar"]/ul/li[5]/a')
directory.click()

driver.implicitly_wait(50) 
country = "Australia" # Adjust this value depending on the country data you want to scrape
country_selection_input = driver.find_element(By.XPATH, '//*[@id="country"]')
select = Select(country_selection_input)
select.select_by_visible_text(country)

city_selection_input = driver.find_element(By.XPATH, '//*[@id="city"]')
select = Select(city_selection_input)
select.select_by_visible_text("All Cities")

search_submit = driver.find_element(By.XPATH, '//*[@id="btn_search"]')
search_submit.click()

# Find the parent div elements
parent_divs = driver.find_elements(By.CLASS_NAME, "groupHQ")

# Initialize lists to store data from each link
company_names = []
all_names = []
all_emails = []
all_phones = []
# Loop through each parent div
for parent_div in parent_divs:
    # Find all child div elements within each parent div
    child_divs = parent_div.find_elements(By.CLASS_NAME, "directory_search_entry")

    # Loop through each child div
    for div in child_divs:
        # Find the <a> tag within the child div
        link = div.find_element(By.TAG_NAME, 'a')

        # Get the original window handle
        original_window = driver.current_window_handle

        # Open the link in a new tab
        link.send_keys(Keys.CONTROL + Keys.RETURN)

        # Switch to the new tab
        new_window = [window for window in driver.window_handles if window != original_window][0]
        driver.switch_to.window(new_window)

        # Add your script for each <a> tag click here
        profile_labels = driver.find_elements(By.CLASS_NAME, "profile_label")
        names = []
        emails = []
        phone = []
        company_name_elements = driver.find_elements(By.CLASS_NAME, "company")
        company_name = company_name_elements[0].text.strip()

        for profile_label_element in profile_labels:
            profile_label_text = profile_label_element.text.strip()
            if profile_label_text == "Name:":
                try:
                    profile_value_element = profile_label_element.find_element(By.XPATH,
                                                                               './following-sibling::div[contains(@class, "profile_val")]')
                    profile_value = profile_value_element.text.strip()
                    names.append(profile_value)
                except NoSuchElementException:
                    print(f"Profile value not found for {profile_label_text}")
            elif profile_label_text == "Email":
                try:
                    profile_value_element = profile_label_element.find_element(By.XPATH,
                                                                               './following-sibling::div[contains(@class, "profile_val")]')
                    profile_value = profile_value_element.text.strip()
                    emails.append(profile_value)
                except NoSuchElementException:
                    print(f"Profile value not found for {profile_label_text}")
            elif profile_label_text == "Email:":
                try:
                    profile_value_element = profile_label_element.find_element(By.XPATH,
                                                                               './following-sibling::div[contains(@class, "profile_val")]')
                    profile_value = profile_value_element.text.strip()
                    emails.append(profile_value)
                except NoSuchElementException:
                    print(f"Profile value not found for {profile_label_text}")
            elif profile_label_text == "Phone:":
                try:
                    profile_value_element = profile_label_element.find_element(By.XPATH,
                                                                               './following-sibling::div[contains(@class, "profile_val")]')
                    profile_value = profile_value_element.text.strip()
                    phone.append(profile_value)
                except NoSuchElementException:
                    print(f"Profile value not found for {profile_label_text}")

        # Join the names and emails lists with appropriate separators
        names_concatenated = ', '.join(names)
        emails_concatenated = '; '.join(emails)
        phones_concatenated = ', '.join(phone)
        # Append data from each link to the lists
        company_names.append(company_name)
        all_names.append(names_concatenated)
        all_emails.append(emails_concatenated)
        all_phones.append(phones_concatenated)
        driver.close()
        time.sleep(10)
        # Switch back to the original tab
        driver.switch_to.window(original_window)

        # Wait a bit before moving to the next link
        time.sleep(10)

    # Re-find the parent div elements after processing all child divs
    parent_divs = driver.find_elements(By.CLASS_NAME, "groupHQ")

# Create a DataFrame with all the collected data
data = {'Company Name': company_names, 'Names': all_names, 'Emails': all_emails, 'Phones': all_phones}
df = pd.DataFrame(data)

# Append the new data to the existing Excel file
try:
    existing_df = pd.read_excel('output1.xlsx')
    combined_df = pd.concat([existing_df, df], ignore_index=True)
    combined_df.to_excel('output1.xlsx', index=False)
except FileNotFoundError:
    df.to_excel(f'{country}.xlsx', index=False)

driver.quit()