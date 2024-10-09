import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

# Initialize WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# URL of the webpage to scrape
url = "https://engineering.nyu.edu/academics/departments/civil-and-urban-engineering/people"

# Open the webpage
driver.get(url)

# Wait for the page to load
time.sleep(5)

# Locate all cards using the common class 'card'
cards = driver.find_elements(By.CLASS_NAME, 'card')

# List to hold all faculty data
faculty_data = []

# Iterate over each card and click on it
for i in range(len(cards)):
    try:
        # Reinitialize the card list to avoid stale element reference
        cards = driver.find_elements(By.CLASS_NAME, 'card')

        # Find the link and click it
        link = cards[i].find_element(By.TAG_NAME, 'a')
        driver.execute_script("arguments[0].scrollIntoView();", link)
        link.click()

        # Wait for the page to load
        time.sleep(3)

        # Dictionary to hold data for the current faculty
        faculty_info = {}

        # Extract the email from the new page
        try:
            email_element = driver.find_element(By.CLASS_NAME, 'field--name-field-email')
            email = email_element.find_element(By.TAG_NAME, 'a').text
            faculty_info["email"] = email
        except:
            faculty_info["email"] = None

        # Extract the name, academic credentials, titles, and department
        try:
            # Extract Name and Academic Credentials
            content_header = driver.find_element(By.CLASS_NAME, 'content-header')
            name = content_header.find_element(By.CLASS_NAME, 'field--name-title').text
            credentials = content_header.find_element(By.CLASS_NAME, 'field--name-field-academic-credential').text
            faculty_info["name"] = name
            faculty_info["academic_credentials"] = credentials

            # Extract Titles
            titles = content_header.find_elements(By.CLASS_NAME, 'field--name-field-title')
            title_list = [title.text for title in titles]
            faculty_info["titles"] = title_list

            # Extract Department
            department_element = content_header.find_element(By.CLASS_NAME, 'field--name-field-department')
            department = department_element.find_element(By.TAG_NAME, 'a').text
            faculty_info["department"] = department

        except Exception as e:
            print(f"Could not extract header information: {e}")

        # Extract Profile Summary
        try:
            summary_element = driver.find_element(By.CLASS_NAME, 'field--name-field-structured-body')
            summary = summary_element.find_element(By.CLASS_NAME, 'field__item').text
            faculty_info["profile_summary"] = summary
        except:
            faculty_info["profile_summary"] = None

        # Extract Research Interests
        try:
            research_element = driver.find_element(By.CLASS_NAME, 'field--name-field-research-interests')
            research_interests = research_element.find_element(By.CLASS_NAME, 'field__item').text
            faculty_info["research_interests"] = research_interests
        except:
            faculty_info["research_interests"] = None

        # Append the faculty data to the list
        faculty_data.append(faculty_info)

        # Navigate back to the previous page
        driver.back()

        # Wait for the page to load again
        time.sleep(3)

    except Exception as e:
        print(f"Could not click on the card: {e}")

# Close the browser
driver.quit()

# Save the faculty data to a JSON file
with open('faculty_data.json', 'w') as json_file:
    json.dump(faculty_data, json_file, indent=4)

print("Data saved to faculty_data.json")
