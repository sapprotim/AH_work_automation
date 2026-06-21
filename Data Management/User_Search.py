import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException

# ========= CONFIG =========
LOGIN_URL = "https://gethealthier.connectedlife.io/"
EMAIL = os.environ.get("GH_EMAIL", "")
PASSWORD = os.environ.get("GH_PASSWORD", "")
OTP_CODE = os.environ.get("GH_OTP", "")
USER_SHEET_PATH = "user_id.xlsx"
USER_COLUMN_NAME = "User Name"
CLICK_FIRST_RESULT = False      # True to click the first matched row
# ==========================

# --- Setup WebDriver ---
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
# chrome_options.add_argument("--headless=new")  # enable if needed
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

def find_otp_input():
    """Locate OTP input; try main DOM then iframes."""
    # try main DOM
    try:
        time.sleep(2)
        return WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((
                By.XPATH,
                "//input[@id='inp' or @name='otp' or @autocomplete='one-time-code' or @placeholder='Enter OTP']"
            ))
        )
    except TimeoutException:
        pass

    # try iframes
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    for frame in frames:
        try:
            driver.switch_to.frame(frame)
            elem = WebDriverWait(driver, 3).until(
                EC.visibility_of_element_located((
                    By.XPATH,
                    "//input[@id='inp' or @name='otp' or @autocomplete='one-time-code' or @placeholder='Enter OTP']"
                ))
            )
            return elem
        except TimeoutException:
            driver.switch_to.default_content()
            continue

    # ensure back to default
    try:
        driver.switch_to.default_content()
    except:
        pass
    return None

def wait_post_login():
    """Wait for a known post-login signal."""
    try:
        WebDriverWait(driver, 25).until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//span[normalize-space()='My Users']")),
                EC.url_contains("#/dashboard"),
                EC.url_contains("#/home")
            )
        )
    except TimeoutException:
        # fallback poll
        for _ in range(50):
            if "#/dashboard" in driver.current_url or "#/home" in driver.current_url:
                return
            if driver.find_elements(By.XPATH, "//span[normalize-space()='My Users']"):
                return
            time.sleep(0.4)

def submit_otp_with_retries(otp_code):
    """
    Fill OTP and submit with several fallbacks, handling staleness.
    """
    attempts = 3
    last_elem = None

    for attempt in range(1, attempts + 1):
        otp_elem = find_otp_input()
        if not otp_elem:
            # maybe auto-submitted or already past OTP
            return

        # type OTP (refetch on staleness)
        try:
            otp_elem.clear()
            otp_elem.send_keys(otp_code)
        except StaleElementReferenceException:
            # refetch once
            otp_elem = find_otp_input()
            if not otp_elem:
                return
            otp_elem.clear()
            otp_elem.send_keys(otp_code)

        # try submit buttons
        submit_xpaths = [
            "//input[@value='Submit']",
            "//button[normalize-space()='Submit']",
            "//button[normalize-space()='Verify']",
            "//button[contains(., 'Submit')]",
            "//button[contains(., 'Verify')]",
            "//input[@type='submit']",
        ]
        clicked = False
        for xp in submit_xpaths:
            try:
                btn = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, xp)))
                btn.click()
                clicked = True
                break
            except Exception:
                continue

        if not clicked:
            # fallback 1: send Enter to active element (avoid stale OTP elem)
            try:
                driver.switch_to.active_element.send_keys(Keys.ENTER)
            except Exception:
                # fallback 2: JS submit (if within a form)
                try:
                    driver.execute_script("""
                        const inp = document.querySelector("input#inp, input[name='otp'], input[autocomplete='one-time-code'], input[placeholder='Enter OTP']");
                        if (inp && inp.form) inp.form.submit();
                    """)
                except Exception:
                    pass

        # wait for either OTP field to go stale, or post-login UI to appear
        try:
            # wait for staleness of the element we just used
            WebDriverWait(driver, 8).until(EC.staleness_of(otp_elem))
            break  # staleness observed -> likely navigated
        except Exception:
            # if not stale, maybe UI already moved on
            if driver.find_elements(By.XPATH, "//span[normalize-space()='My Users']") or \
               "#/dashboard" in driver.current_url or "#/home" in driver.current_url:
                break
            # one more try if attempts left
            if attempt < attempts:
                time.sleep(1.5)
                continue

    # make sure we are out of any frame
    try:
        driver.switch_to.default_content()
    except:
        pass

try:
    # --- Login ---
    driver.get(LOGIN_URL)
    wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@name='username']"))).send_keys(EMAIL)
    driver.find_element(By.XPATH, "//input[@name='password']").send_keys(PASSWORD)
    driver.find_element(By.XPATH, "//input[@value='Sign In']").click()

    # --- OTP step with robust submit ---
    submit_otp_with_retries(OTP_CODE)

    # --- Wait for app load after OTP ---
    wait_post_login()

    # --- Go to My Users (if visible) ---
    try:
        driver.find_element(By.XPATH, "//span[normalize-space()='My Users']").click()
    except Exception:
        pass
    time.sleep(1)

    # --- Load users from Excel ---
    df = pd.read_excel(USER_SHEET_PATH)
    if USER_COLUMN_NAME not in df.columns:
        raise ValueError(f"'{USER_COLUMN_NAME}' column not found in {USER_SHEET_PATH}")

    users = (
        df[USER_COLUMN_NAME]
        .dropna()
        .astype(str)
        .map(str.strip)
        .tolist()
    )

    # --- Search loop ---
    search_box_xpath = "//input[@placeholder='Search']"
    for u in users:
        if not u:
            continue

        # ensure search box visible; refresh once if needed
        try:
            box = wait.until(EC.visibility_of_element_located((By.XPATH, search_box_xpath)))
        except TimeoutException:
            driver.refresh()
            time.sleep(2)
            box = wait.until(EC.visibility_of_element_located((By.XPATH, search_box_xpath)))

        box.clear()
        box.send_keys(u)
        time.sleep(1.5)  # wait for table update

        rows = driver.find_elements(By.XPATH, "//div[@class='table-wrap']//table/tbody/tr")
        if rows and CLICK_FIRST_RESULT:
            try:
                rows[0].find_element(By.XPATH, ".//td[1]//span").click()
            except Exception:
                pass

        print(f"[{'FOUND' if rows else 'NOT FOUND'}] {u}")

        # clear for next search
        try:
            driver.find_element(By.XPATH, search_box_xpath).clear()
        except Exception:
            pass
        time.sleep(0.5)

    print("Done.")
finally:
    driver.quit()
