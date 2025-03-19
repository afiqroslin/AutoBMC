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
options.add_experimental_option("detach", True)  # detach chrome from program after open, to keep it running after script finished execution
options.add_experimental_option('excludeSwitches', ['enable-logging'])
caps = options.to_capabilities()
caps["acceptInsecureCerts"] = True  # proceed to unsafe if safety alert popped up


def AutoBMC(BMC_IP, USERNAME, USER, PWD):
    print(f"bmc: {BMC_IP}")
    print(f"username {USERNAME}")

    print("Redirecting...")
    driver = webdriver.Chrome(options=options)
    driver.get("https://" + BMC_IP) #Get server IP address/domain name

    # Insert username
    wait = WebDriverWait(driver, 20)
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div[2]/div/div/div[1]/div[1]/div/section/main/article/div[2]/div[1]/div[2]/div/form/div[1]/div[1]/div/label/input"))).send_keys(USERNAME)

    # Insert password fill and click login
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div[2]/div/div/div[1]/div[1]/div/section/main/article/div[2]/div[1]/div[2]/div/form/div[1]/div[2]/div/label/input"))).send_keys(PWD)

    #Click login
    wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div/div[2]/div/div/div[1]/div[1]/div/section/main/article/div[2]/div[1]/div[2]/div/form/div[1]/div[3]/button"))).click()


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
