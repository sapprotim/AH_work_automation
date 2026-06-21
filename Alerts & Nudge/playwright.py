import os
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError

EMAIL = os.environ.get("GH_EMAIL", "")
PASSWORD = os.environ.get("GH_PASSWORD", "")
OTP = os.environ.get("GH_OTP", "")

# Read Excel files
df_user_1 = pd.read_excel('user list.xlsx')
df_alerts_1 = pd.read_excel('Alerts Nudge list.xlsx')

df_user_2 = pd.read_excel('user list.xlsx')
df_alerts_2 = pd.read_excel('Alerts Nudge list.xlsx')

df_user_3 = pd.read_excel('user list.xlsx')
df_alerts_3 = pd.read_excel('Alerts Nudge list.xlsx')

usernames_1 = df_user_1['ALERTS ONLY FOR STUDY ARM 3'].dropna().tolist()
nudge_alerts_1 = df_alerts_1['ALERTS ONLY FOR STUDY ARM 3'].dropna().tolist()

usernames_2 = df_user_2['NUDGES ONLY FOR STUDY ARM 2 AND 3 (English)'].dropna().tolist()
nudge_alerts_2 = df_alerts_2['NUDGES ONLY FOR STUDY ARM 2 AND 3 (English)'].dropna().tolist()

usernames_3 = df_user_3['NUDGES ONLY FOR STUDY ARM 2 AND 3 (Chinese)'].dropna().tolist()
nudge_alerts_3 = df_alerts_3['NUDGES ONLY FOR STUDY ARM 2 AND 3 (Chinese)'].dropna().tolist()

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()

    # Go to site
    page.goto("https://gethealthier-dev.connectedlife.io/", timeout=60000)

    # Login
    page.locator("//div[@class='login-form']//div[1]//label[1]").fill(EMAIL)
    page.locator("//div[@class='divisions']//div[2]//label[1]//input[1]").fill(PASSWORD)
    page.locator("//input[@value='Sign In']").click()
    page.wait_for_load_state('networkidle')

    # OTP input
    page.locator("#inp").fill(OTP)
    page.locator("//input[@value='Submit']").click()
    page.wait_for_load_state('networkidle')

    # Navigate Programme Management -> Alerts/Nudges
    try:
        page.locator("//span[normalize-space()='Programme Management']").click()
        page.locator("//a[normalize-space()='Alerts/Nudges']").click()
    except TimeoutError:
        page.reload()
        page.locator("//span[normalize-space()='Programme Management']").click()
        page.locator("//a[normalize-space()='Alerts/Nudges']").click()

    # Assign Study Arm 3 Alerts
    for nudge_alert in nudge_alerts_1:
        try:
            page.locator("//input[@placeholder='Search']").fill(nudge_alert)
        except TimeoutError:
            page.locator("//img[@title='Refresh']").click()
            page.locator("//input[@placeholder='Search']").fill(nudge_alert)

        page.wait_for_timeout(2000)
        page.locator("//tbody/tr[1]/td[7]/div[1]/img[2]").click()
        page.wait_for_timeout(2000)

        for user in usernames_1:
            try:
                user_str = str(int(user))
                search_box = page.locator("//input[@placeholder='Search']")
                search_box.fill(user_str)
                page.wait_for_timeout(2000)
                page.locator("//table[1]/tbody[1]/tr[1]/td[1]/label[1]/span[1]").click()
                page.wait_for_timeout(1000)
                search_box.fill("")
            except Exception as e:
                print(f"User {user_str} assign failed: {e}")

        page.locator("//div[@class='add-patient-btn ng-star-inserted']").click()
        page.wait_for_timeout(5000)
        page.locator("//input[@placeholder='Search']").fill("")
        page.locator("//img[@title='Refresh']").click()
        page.wait_for_timeout(1000)
        print(f"Assigned {nudge_alert} to users")

    # Repeat for Study Arm 2 and 3 (English) Nudges
    for nudge_alert in nudge_alerts_2:
        try:
            page.locator("//input[@placeholder='Search']").fill(nudge_alert)
        except TimeoutError:
            page.locator("//img[@title='Refresh']").click()
            page.locator("//input[@placeholder='Search']").fill(nudge_alert)

        page.wait_for_timeout(2000)
        page.locator("//tbody/tr[1]/td[7]/div[1]/img[2]").click()
        page.wait_for_timeout(2000)

        for user in usernames_2:
            try:
                user_str = str(int(user))
                search_box = page.locator("//input[@placeholder='Search']")
                search_box.fill(user_str)
                page.wait_for_timeout(2000)
                page.locator("//table[1]/tbody[1]/tr[1]/td[1]/label[1]/span[1]").click()
                page.wait_for_timeout(1000)
                search_box.fill("")
            except Exception as e:
                print(f"User {user_str} assign failed: {e}")

        page.locator("//div[@class='add-patient-btn ng-star-inserted']").click()
        page.wait_for_timeout(5000)
        page.locator("//input[@placeholder='Search']").fill("")
        page.locator("//img[@title='Refresh']").click()
        page.wait_for_timeout(1000)
        print(f"Assigned {nudge_alert} to users")

    # Repeat for Study Arm 2 and 3 (Chinese) Nudges
    for nudge_alert in nudge_alerts_3:
        try:
            page.locator("//input[@placeholder='Search']").fill(nudge_alert)
        except TimeoutError:
            page.locator("//img[@title='Refresh']").click()
            page.locator("//input[@placeholder='Search']").fill(nudge_alert)

        page.wait_for_timeout(2000)
        page.locator("//tbody/tr[1]/td[7]/div[1]/img[2]").click()
        page.wait_for_timeout(2000)

        for user in usernames_3:
            try:
                user_str = str(int(user))
                search_box = page.locator("//input[@placeholder='Search']")
                search_box.fill(user_str)
                page.wait_for_timeout(2000)
                page.locator("//table[1]/tbody[1]/tr[1]/td[1]/label[1]/span[1]").click()
                page.wait_for_timeout(1000)
                search_box.fill("")
            except Exception as e:
                print(f"User {user_str} assign failed: {e}")

        page.locator("//div[@class='add-patient-btn ng-star-inserted']").click()
        page.wait_for_timeout(5000)
        page.locator("//input[@placeholder='Search']").fill("")
        page.locator("//img[@title='Refresh']").click()
        page.wait_for_timeout(1000)
        print(f"Assigned {nudge_alert} to users")

    browser.close()

