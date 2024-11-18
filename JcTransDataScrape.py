import time
import openpyxl
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains

# Setup Chrome options
options = Options()
options.add_experimental_option("detach", True)  # Keeps the window open

# Path to ChromeDriver
chrome_driver_path = 'C:\\chromedriver-win64\\chromedriver.exe'  # Replace with the actual path

# Setup WebDriver
driver = webdriver.Chrome(service=Service(chrome_driver_path), options=options)

driver.get("https://www.jctrans.com/en")
driver.maximize_window()

# Wait for the page to load before locating elements
driver.implicitly_wait(50)

# Click the link that opens the login form
x_path_expression = '//*[@id="app"]/div[1]/div[1]/div[1]/div[2]/button'
link = driver.find_element(By.XPATH, x_path_expression)
link.click()

# Fill in the username and password inputs
username_input = driver.find_element(By.XPATH, '//*[@id="pane-first"]/form/div[1]/div/div[1]/div/div/div/input')
username_input.send_keys("")  # Add your username

password_input = driver.find_element(By.XPATH, '//*[@id="pane-first"]/form/div[2]/div/div/div/div/div/input')
password_input.send_keys("")  # Add your password

# Click the submit button (login button)
submit = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[2]/div[3]/div[3]/div[3]/button')
submit.click()

input("Please complete the CAPTCHA or verification manually and then press Enter to continue...")

driver.implicitly_wait(50)

country_name = "Dubai"
max_pages = 101  # Set your desired maximum number of pages here

# Initialize a list to store all data
all_data = []

page_counter = 0

while page_counter < max_pages:
    print(f"Processing page {page_counter + 1}...")

    # Find all li elements under the specified ul
    li_elements = driver.find_elements(By.XPATH, '//*[@id="app"]/div[1]/section/div/div/div[1]/div[3]/div[2]/div/div/ul/li')

    for li in li_elements:
        # Click on each li element
        cargo_name = li.find_element(By.CLASS_NAME, "membership-list-content-center-list-item-left").text  # Extract the cargo name
        li.find_element(By.CLASS_NAME, "membership-list-content-center-list-item-left").click()

        # Switch to the new tab opened
        driver.switch_to.window(driver.window_handles[-1])

        # Initialize strings to store the aggregated information
        all_names = []
        all_emails = []
        all_phones = []

        # Find all contact cards
        contact_cards = driver.find_elements(By.CLASS_NAME, "contactCard")

        for card in contact_cards:
            try:
                # Extract the name
                name = card.find_element(By.CLASS_NAME, "font-700").text
                all_names.append(name)

                # Extract the email
                email_element = card.find_element(By.XPATH,
                                                  ".//div[@class='flex items-center justify-start bg-[#ECF0F6] rounded-[4px] h-30px pr-10px pl-10px mr-5px mb-10px']/div[@class='content pl-5px']")
                email = email_element.text
                all_emails.append(email)

                # Extract the phone number (if it exists)
                phone_elements = card.find_elements(By.XPATH,
                                                    ".//div[@class='flex items-center justify-start bg-[#ECF0F6] rounded-[4px] h-30px pr-10px pl-10px mr-5px mb-10px']/div[@class='content pl-5px']")
                phone_number = ""
                if len(phone_elements) > 1:
                    phone_number = phone_elements[
                        -1].text  # Assuming the phone number is always the last element in this div
                all_phones.append(phone_number)

            except NoSuchElementException as e:
                print(f"An element was not found: {e}")
                continue

        # Aggregate data for the cargo company
        all_data.append({
            'Country': country_name,
            'Cargo Name': cargo_name,
            'Names': ", ".join(all_names),
            'Emails': "; ".join(all_emails),
            'Phones': ", ".join(all_phones)
        })

        # CLOSE THE CONTACT US TAB AND SWITCH BACK
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

    # Check if there is a next button to click
    try:
        next_button = driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/section/div/div/div[1]/div[3]/div[2]/div/div/div/button[2]')
        next_button.click()
        page_counter += 1
        time.sleep(3)  # Wait for the next page to load
    except NoSuchElementException:
        print("No more pages available or unable to find the next button.")
        break

# Convert the collected data into a DataFrame
df = pd.DataFrame(all_data)

# Save the DataFrame to an Excel file
df.to_excel(f"contact_details_{country_name}.xlsx", index=False)
print(f"Contact details saved to contact_details_{country_name}.xlsx")

# Close the driver
driver.quit()
