import os
import time
import pandas as pd
from playwright.sync_api import sync_playwright

EMAIL = os.environ.get("GH_EMAIL", "")
PASSWORD = os.environ.get("GH_PASSWORD", "")
OTP = os.environ.get("GH_OTP", "")


def assign_alerts_nudges(page, usernames, nudge_alerts):
    for nudge_alert in nudge_alerts:
        try:
            page.locator("//input[@placeholder='Search']").fill(nudge_alert)
        except:
            time.sleep(5)
            page.locator("//img[@title='Refresh']").click()
            time.sleep(5)
            page.locator("//input[@placeholder='Search']").fill(nudge_alert)
        time.sleep(2)
        page.locator("//tbody/tr[1]/td[7]/div[1]/img[2]").click()
        time.sleep(2)
        for user in usernames:
            user = int(user)
            try:
                page.locator("//input[@placeholder='Search']").fill(str(user))
            except:
                time.sleep(5)
                page.locator("//input[@placeholder='Search']").clear()
                page.locator("//input[@placeholder='Search']").fill(str(user))
            time.sleep(4)
            page.locator("//div[@class='table-wrap']//table/tbody/tr/td[1]//span").click()
            time.sleep(1)
            page.locator("//input[@placeholder='Search']").clear()
            print(user)
        time.sleep(1)
        page.locator("//div[@class='add-patient-btn ng-star-inserted']").click()
        time.sleep(7)
        page.locator("//input[@placeholder='Search']").clear()
        page.locator("//img[@title='Refresh']").click()
        time.sleep(1)
        print(nudge_alert)
        time.sleep(2)


with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)  # headless mode enabled
    page = browser.new_page()

    page.goto("https://gethealthier.connectedlife.io/#/login")
    page.wait_for_timeout(1000)

    # Login
    page.locator("//div[@class='login-form']//div[1]//label[1]").fill(EMAIL)
    page.locator("//div[@class='divisions']//div[2]//label[1]//input[1]").fill(PASSWORD)
    page.locator("//input[@value='Sign In']").click()
    time.sleep(5)

    # OTP Input
    page.locator("//input[@id='inp']").fill(OTP)
    page.locator("//input[@value='Submit']").click()
    time.sleep(5)

    # Navigate to Programme Management -> Alerts/Nudges
    try:
        page.locator("//span[normalize-space()='Programme Management']").click()
        time.sleep(2)
        page.locator("//a[normalize-space()='Alerts/Nudges']").click()
        time.sleep(2)
    except:
        page.reload()
        time.sleep(2)
        page.locator("//span[normalize-space()='Programme Management']").click()
        time.sleep(2)
        page.locator("//a[normalize-space()='Alerts/Nudges']").click()
        time.sleep(2)

    # Read Excel files and assign alerts/nudges
    user_list_file = 'user list.xlsx'
    alert_nudge_list_file = 'Alerts Nudge list.xlsx'

    df_user_1 = pd.read_excel(user_list_file)
    usernames_1 = df_user_1['ALERTS ONLY FOR STUDY ARM 3'].dropna().tolist()
    df_alerts_1 = pd.read_excel(alert_nudge_list_file)
    nudge_alerts_1 = df_alerts_1['ALERTS ONLY FOR STUDY ARM 3'].dropna().tolist()
    assign_alerts_nudges(page, usernames=usernames_1, nudge_alerts=nudge_alerts_1)

    df_user_2 = pd.read_excel(user_list_file)
    usernames_2 = df_user_2['NUDGES ONLY FOR STUDY ARM 2 AND 3 (English)'].dropna().tolist()
    df_alerts_2 = pd.read_excel(alert_nudge_list_file)
    nudge_alerts_2 = df_alerts_2['NUDGES ONLY FOR STUDY ARM 2 AND 3 (English)'].dropna().tolist()
    assign_alerts_nudges(page, usernames=usernames_2, nudge_alerts=nudge_alerts_2)

    df_user_3 = pd.read_excel(user_list_file)
    usernames_3 = df_user_3['NUDGES ONLY FOR STUDY ARM 2 AND 3 (Chinese)'].dropna().tolist()
    df_alerts_3 = pd.read_excel(alert_nudge_list_file)
    nudge_alerts_3 = df_alerts_3['NUDGES ONLY FOR STUDY ARM 2 AND 3 (Chinese)'].dropna().tolist()
    assign_alerts_nudges(page, usernames=usernames_3, nudge_alerts=nudge_alerts_3)

    browser.close()
