import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By

EMAIL = os.environ.get("GH_EMAIL", "")
PASSWORD = os.environ.get("GH_PASSWORD", "")
OTP = os.environ.get("GH_OTP", "")

# Setup Selenium WebDriver
driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://gethealthier-dev.connectedlife.io/")
driver.implicitly_wait(100)

# Log in
driver.find_element(By.XPATH, "//div[@class='login-form']//div[1]//label[1]").send_keys(EMAIL)
driver.find_element(By.XPATH, "//div[@class='divisions']//div[2]//label[1]//input[1]").send_keys(PASSWORD)
driver.find_element(By.XPATH, "//input[@value='Sign In']").click()
time.sleep(5)  # Allow time for page to load

# OTP Input
otp = OTP
driver.find_element(By.XPATH, "//input[@id='inp']").send_keys(otp)
driver.find_element(By.XPATH, "//input[@value='Submit']").click()
time.sleep(5)  # Allow time for submission and page navigation

# Navigate to Programme Management -> Alerts/Nudges
try:
    driver.find_element(By.XPATH, "//span[normalize-space()='Programme Management']").click()
    time.sleep(2)  # Wait for the menu to appear
    driver.find_element(By.XPATH, "//a[normalize-space()='Alerts/Nudges']").click()
    time.sleep(2)  # Wait for Alerts/Nudges page to load
except:
    driver.refresh()
    driver.find_element(By.XPATH, "//span[normalize-space()='Programme Management']").click()
    time.sleep(2)  # Wait for the menu to appear
    driver.find_element(By.XPATH, "//a[normalize-space()='Alerts/Nudges']").click()
    time.sleep(2)  # Wait for Alerts/Nudges page to load
# Define function for searching and assigning nudges/alerts
def assign_alerts_nudges(driver, usernames, nudge_alerts):
    for nudge_alert in nudge_alerts:
        try:
            driver.find_element(By.XPATH, "//input[@placeholder='Search']").send_keys(nudge_alert)
        except:
            time.sleep(5)
            driver.find_element(By.XPATH, "//img[@title='Refresh']").click()
            time.sleep(5)
            driver.find_element(By.XPATH, "//input[@placeholder='Search']").send_keys(nudge_alert)
        time.sleep(2)
        driver.find_element(By.XPATH, "//tbody/tr[1]/td[7]/div[1]/img[2]").click()
        time.sleep(2)
        for user in usernames:
            user =  int(user)
            try:
                driver.find_element(By.XPATH, "//input[@placeholder='Search']").send_keys(user)
            except:
                time.sleep(5)
                driver.find_element(By.XPATH, "//input[@placeholder='Search']").clear()
                driver.find_element(By.XPATH, "//input[@placeholder='Search']").send_keys(user)
            time.sleep(2)
            driver.find_element(By.XPATH, "/html[1]/body[1]/app-root[1]/div[3]/app-multiple-user-alerts[1]/div[1]/app-doctor-assign-patients-list[1]/div[1]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/label[1]/span[1]").click()
            time.sleep(1)
            driver.find_element(By.XPATH, "//input[@placeholder='Search']").clear()
        time.sleep(1)
        driver.find_element(By.XPATH, "//div[@class='add-patient-btn ng-star-inserted']").click()
        time.sleep(7)
        driver.find_element(By.XPATH, "//input[@placeholder='Search']").clear()
        driver.find_element(By.XPATH, "//img[@title='Refresh']").click()
        time.sleep(1)
        print(nudge_alert, user)
        time.sleep(2)


# Read user and alert/nudge lists from Excel
user_list_file = 'user list.xlsx'
alert_nudge_list_file = 'Alerts Nudge list.xlsx'

# Study Arm 3 Alerts and # Assign Study Arm 3 Alerts
df_user_1 = pd.read_excel(user_list_file)
usernames_1 = df_user_1['ALERTS ONLY FOR STUDY ARM 3'].dropna().tolist()
df_alerts_1 = pd.read_excel(alert_nudge_list_file)
nudge_alerts_1 = df_alerts_1['ALERTS ONLY FOR STUDY ARM 3'].dropna().tolist()
assign_alerts_nudges(driver, usernames=usernames_1, nudge_alerts = nudge_alerts_1)

# Study Arm 2 and 3 (English) Nudges and # Assign Study Arm 2 and 3 (English) Nudges
df_user_2 = pd.read_excel(user_list_file)
usernames_2 = df_user_2['NUDGES ONLY FOR STUDY ARM 2 AND 3 (English)'].dropna().tolist()
df_alerts_2 = pd.read_excel(alert_nudge_list_file)
nudge_alerts_2 = df_alerts_2['NUDGES ONLY FOR STUDY ARM 2 AND 3 (English)'].dropna().tolist()
assign_alerts_nudges(driver, usernames=usernames_2, nudge_alerts =nudge_alerts_2)

# Study Arm 2 and 3 (Chinese) Nudges and # Assign Study Arm 2 and 3 (Chinese) Nudges
df_user_3 = pd.read_excel(user_list_file)
usernames_3 = df_user_3['NUDGES ONLY FOR STUDY ARM 2 AND 3 (Chinese)'].dropna().tolist()
df_alerts_3 = pd.read_excel(alert_nudge_list_file)
nudge_alerts_3 = df_alerts_3['NUDGES ONLY FOR STUDY ARM 2 AND 3 (Chinese)'].dropna().tolist()
assign_alerts_nudges(driver, usernames =usernames_3, nudge_alerts = nudge_alerts_3)

# Cleanup and exit
driver.quit()

