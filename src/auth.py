from easygui import *
import logging
from selenium import webdriver
#from src.config import AppConfiguration     #src.config
#ImportError: cannot import name 'AppConfiguration' from partially initialized module 'src.config' (most likely due to a circular import) (C:\entryBot\entrybot-20210109T1735\src\config.py)

import os
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException



def enterlogincredentials():
    if config["driver"] == "firefox":
        path = cwd + '/drivers/firefoxdriver.exe'
        driver = webdriver.Firefox(executable_path=path)
    else:
        path = cwd + '/drivers/chromedriver.exe'
        driver = webdriver.Chrome(executable_path=path)
    driver.get(config["URL"])
    assert "ICMR" in driver.title
    driver.find_element_by_name("username").send_keys(config["u_name"])
    driver.find_element_by_name("passwd").send_keys(config["pwd"])
    logging.info("ICMR Page loaded at  " + datetime.now().strftime('%d_%m_%Y__%H_%M_%S_%f'))

def askforlogincred():
    u_name = textbox('Type in the User Name for ICMR COVID Portal', 'ICMR Website User Name ', "", 0).strip
    pwd = passwordbox('Type in the Password for ICMR COVID Portal', 'ICMR Website Password', "", 0)
    appConfig.setConfig("u_name", u_name)
    appConfig.setConfig("pwd", pwd)
    print("User ID and Password are collected")
    logging.info("User ID and Password are collected")

def clickonlogin(driver):
    try:
        driver.find_element_by_id("login_btn").click()
        driver.implicitly_wait(2)
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        logging.info("Login Failed, please check login credentials ")
        logging.info(alert.text + str(alert))
        print("Login Failed, please check login credentials ")
        return False
    except TimeoutException:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "patient_id"))
        )
        return True
