import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException

# Set up proxy
# proxy = "13.126.79.133:3128"
chrome_options = webdriver.ChromeOptions() 
# chrome_options.add_argument(f'--proxy-server={proxy}')

# Set the path to your ChromeDriver
service = Service("C:/Users/Shekhani Laptop/Downloads/New folder (25)/chromedriver-win64/chromedriver.exe")

# Initialize the Chrome driver with proxy settings
driver = webdriver.Chrome(service=service, options=chrome_options)

# Set an implicit wait
driver.implicitly_wait(50)

# Navigate to the website
driver.get("https://safer.fmcsa.dot.gov/CompanySnapshot.aspx")

# Wait for the form and elements to be present with increased explicit wait time
wait = WebDriverWait(driver, 50)

# Create an empty list to store data
all_data = []

# Helper function to save data to Excel periodically
def save_data_to_excel(data, filename="scraped_data_with_emails.xlsx"):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"Data saved to {filename}")

# Loop through numbers from 100200 to 100231
for number in range(1322222, 1400000): 
    try:
        # Select MC/MX option
        mc_mx_radio_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='MC_MX']")))
        mc_mx_radio_button.click()

        # Input the MC/MX number
        input_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='query_string']")))
        input_field.clear()
        input_field.send_keys(str(number))

        # Click the search button
        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/form/p/table/tbody/tr[4]/td/input")))
        search_button.click()

        # Handle cases with no records or inactive records
        if "No records matching" in driver.page_source:
            print(f"No record found for MC/MX Number = {number}")
            driver.get("https://safer.fmcsa.dot.gov/CompanySnapshot.aspx")
            continue
        if "Record Inactive" in driver.page_source:
            print(f"Record Inactive for MC/MX Number = {number}")
            driver.get("https://safer.fmcsa.dot.gov/CompanySnapshot.aspx")
            continue 

        field_data = {}
        fields = {
            "Operating Authority Status": "/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[8]/td",
            "USDOT Status": "/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[4]/td[1]",
            "Entity Type": "/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[3]/td", 
            "MC/MX/FF Number(s)": "/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[9]/td/a",
            "Legal Name": "/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[11]/td",
            "Phone Number": "/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[14]/td",
            "Mailing Address": "/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/center[1]/table/tbody/tr[15]/td"
        }

        # Extract data for the current record
        for field_name, xpath in fields.items():
            try:
                element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                field_data[field_name] = element.text
            except TimeoutException:
                print(f"Could not extract {field_name} for MC/MX Number = {number}")
                continue

        # Check conditions to save the record
        if (field_data.get("Entity Type", "").strip() == "CARRIER" and 
            "ACTIVE" in field_data.get("USDOT Status", "").strip() and 
            "AUTHORIZED FOR Property" in field_data.get("Operating Authority Status", "").strip()):

            filtered_data = {
                "MC/MX/FF Number(s)": field_data.get("MC/MX/FF Number(s)", ""),
                "Legal Name": field_data.get("Legal Name", ""),
                "Phone Number": field_data.get("Phone Number", ""),
                "Mailing Address": field_data.get("Mailing Address", "")
            }

            # additional link to get email
            try:
                first_link = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/p/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table[1]/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/font/a")))
                first_link.click()

                additional_link = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[2]/article/div[2]/div[2]/section/a[1]")))
                additional_link.click()

                # Extract email address
                email_element = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[8]/div/div/article/div/div/ul[1]/li[7]/span")))
                filtered_data["Email"] = email_element.text
            except TimeoutException:
                print(f"Could not extract email for MC/MX Number = {number}")

            # Append the filtered data to the all_data list
            all_data.append(filtered_data)

        # Save the data after every few iterations
        if number % 10 == 0:
            save_data_to_excel(all_data)

        # Go back to the search page for the next number
        driver.get("https://safer.fmcsa.dot.gov/CompanySnapshot.aspx")

    except TimeoutException:
        print(f"Timeout while processing MC/MX Number = {number}. Skipping...")
        driver.get("https://safer.fmcsa.dot.gov/CompanySnapshot.aspx")
        continue

# Save the final data at the end
save_data_to_excel(all_data)

print("Script completed. Data saved to 'scraped_data_with_emails.xlsx'.")

# Close the browser
# driver.quit()
# Keep the browser open by waiting for user input
input("Press Enter to close the browser...")
