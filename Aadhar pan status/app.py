import os
import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Load the Excel file
input_file = "data.xlsx"  # Input Excel file name
output_file = "E:\\Nandhini\\Aadhar pan status\\output.xlsx"  # Output Excel file path
data = pd.read_excel(input_file, dtype={"PAN": str})  # Treat PAN as string, but leave Aadhaar as numeric

# Add a column for the result
data["Link Status"] = ""  # We only need the link status

# Set Chrome options to bypass bot detection
chrome_options = Options()
chrome_options.add_argument("--start-maximized")  # Maximize browser window
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("useAutomationExtension", False)
chrome_options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36"
)

# Initialize the WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Bypass bot detection
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined
        });
    """
})

# Open the Income Tax e-Filing portal
driver.get("https://www.incometax.gov.in/iec/foportal")
time.sleep(random.uniform(3, 5))  # Add realistic delay

# Navigate to Aadhaar-PAN Link Status page once
link_status_page = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.LINK_TEXT, "Link Aadhaar Status"))
)
link_status_page.click()
time.sleep(random.uniform(2, 4))  # Random delay

# Loop through each row in the Excel file
for index, row in data.iterrows():
    while True:  # Retry loop for each PAN
        try:
            print(f"Checking Aadhaar-PAN status for PAN: {row['PAN']}")

            # Step 2: Locate input fields for PAN and Aadhaar
            pan_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "mat-input-element"))  # Class name for PAN input field
            )
            aadhaar_input = driver.find_element(By.XPATH, '//*[@id="mat-input-1"]')  # XPath for Aadhaar input field

            # Clear and input PAN and Aadhaar numbers
            pan_input.clear()
            aadhaar_input.clear()
            pan_input.send_keys(row["PAN"])
            aadhaar_input.send_keys(str(row["Aadhaar"]))  # Ensure Aadhaar is sent as a number, not string

            # Locate and click the submit button
            submit_button = driver.find_element(By.CLASS_NAME, "large-button-primary")  # Class name for submit button
            submit_button.click()
            time.sleep(random.uniform(3, 6))  # Wait for result

            # Check for success modal
            try:
                success_modal = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "successScreen"))  # Modal ID for success
                )
                result_message = driver.find_element(By.ID, "linkAadhaarSuccess_desc").text  # Extract success message
                print(f"Success result message for PAN {row['PAN']}: {result_message}")
                data.at[index, "Link Status"] = result_message  # Store the result message in the DataFrame

                # Close the success modal
                close_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "linkAadhaarSuccessClose"))  # Close button ID for success
                )
                close_button.click()
                time.sleep(3)  # Close button wait time set to 3 seconds

                # Break the retry loop after successful result
                break

            except:
                # If success modal is not found, check for failure modal
                failure_modal = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, "failure"))  # Modal ID for failure
                )
                result_message = driver.find_element(By.ID, "linkAadhaarFailure_desc").text  # Extract failure message
                print(f"Failure result message for PAN {row['PAN']}: {result_message}")
                data.at[index, "Link Status"] = result_message  # Store the failure message in the DataFrame

                # Close the failure modal
                close_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "linkAadhaarFailureClose"))  # Close button ID for failure
                )
                close_button.click()
                time.sleep(3)  # Close button wait time set to 3 seconds

                # Break the retry loop after failure result
                break

            # Refresh the page for the next iteration
            driver.refresh()
            # No 5-second delay after refresh now

        except Exception as e:
            print(f"Error processing PAN {row['PAN']}: {e}")
            data.at[index, "Link Status"] = "Error: Could not fetch status"
            time.sleep(random.uniform(3, 5))  # Add a small delay before retrying the PAN

# Close the browser
driver.quit()

# Ensure the directory exists
os.makedirs(os.path.dirname(output_file), exist_ok=True)

# Save the updated Excel file
data.to_excel(output_file, index=False)
print(f"Results saved to {output_file}")
