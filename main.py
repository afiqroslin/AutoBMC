# Project: Automate BMC Systems Login
# Author: Afiq Roslin
# GitHub: Afiq Roslin

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import time

path = r"C:\Users\afiq0\PycharmProjects\AutoBMC\chromedriver_win32\chromedriver.exe"  # path to chromedriver program
options = webdriver.ChromeOptions()  # Chrome WebDriver API
options.add_experimental_option("detach", True)  # detach chrome from program after open
options.add_experimental_option('excludeSwitches', ['enable-logging'])
caps = options.to_capabilities()
caps["acceptInsecureCerts"] = True  # proceed to unsafe if safety alert popped up


def AutoBMC(BMC_IP, USERNAME, USER, PWD):
    print(f"bmc: {BMC_IP}")
    print(f"username {USERNAME}")

    print("Redirecting...")
    driver = webdriver.Chrome(options=options)
    driver.get("http://" + BMC_IP) #Get server IP address/domain name

    # Insert username and click next
    wait = WebDriverWait(driver, 20)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@autocomplete='username']"))).send_keys(USERNAME)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='button'][contains(.,'Next')]"))).click()

    # Twitter will detect automation as suspicious activity, so need to bypass by providing the info needed
    wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@autocomplete='on']"))).send_keys(USER)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='button'][contains(.,'Next')]"))).click()

    # Insert password fill and click login
    wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@autocomplete='current-password']"))).send_keys(PWD)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='button'][contains(.,'Log in')]"))).click()


if __name__ == '__main__':
    while True:
        df = pd.read_excel(r'C:\Users\afiq0\PycharmProjects\AutoBMC\autobmc.xlsx')  # read excel file contains hostname
        print(df.iloc[1:, 0])
        print('\nChoose Hostname\n      or\nPress enter to exit..')  # Show name to select, usually each servers has its own hostname
        userInput = int(input())

        if not userInput:
            continue

        # Take input (ip address, username, password) from the excel file
        if userInput in df.index:
            AutoBMC(BMC_IP=df.iloc[userInput].BMC_IP, USERNAME=df.iloc[userInput].Username,
                    USER=df.iloc[userInput].user, PWD=df.iloc[userInput].Password)

        else:
            continue
