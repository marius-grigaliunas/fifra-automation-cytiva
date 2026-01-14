
# login_test.py

import time
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains


# -----------------------------
LOGIN_URL = "https://pallprod.enlabel.com/Login.aspx?ReturnUrl=%2f"  
USERNAME = "mariusg"                        
PASSWORD = "LietuvaTevyneMusu!123"                        
# Choose the locator strategy that matches your page.
# Replace the tuples below with the real IDs/NAMEs/XPaths/CSS of your fields.

USERNAME_LOCATOR = (By.ID, "ctl00_ContentPlaceHolder1_txtUserName")        
PASSWORD_LOCATOR = (By.ID, "ctl00_ContentPlaceHolder1_txtPassword")        
SUBMIT_LOCATOR   = (By.ID, "ctl00_ContentPlaceHolder1_btnLogin") 

LOT_NUMBER = "UE4376"


def _wait_ready_and_ajax(driver, timeout=30):
    """Wait for document.readyState == 'complete' and for jQuery to be idle (if present)."""
    w = WebDriverWait(driver, timeout)
    w.until(lambda d: d.execute_script("return document.readyState") == "complete")
    try:
        w.until(lambda d: d.execute_script("return (window.jQuery ? jQuery.active : 0) === 0"))
    except Exception:
        pass  # jQuery not present on all pages

def _switch_into_frame_if_needed(driver, locator, probe_timeout=2):
    """
    Ensure Selenium is in the DOM context that contains `locator`.
    Try default content first; otherwise iterate top-level iframes.
    Returns True if found (and switched if needed), else False.
    """
    by, value = locator
    driver.switch_to.default_content()
    try:
        WebDriverWait(driver, probe_timeout).until(EC.presence_of_element_located(locator))
        return True
    except TimeoutException:
        pass

    for fr in driver.find_elements(By.TAG_NAME, "iframe"):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(fr)
            WebDriverWait(driver, probe_timeout).until(EC.presence_of_element_located(locator))
            return True
        except (TimeoutException, StaleElementReferenceException):
            continue

    driver.switch_to.default_content()
    return False

def navigate_after_login(driver, wait):
    """Navigate to ManageDatabases, open the second row entry, and set filter dropdowns."""
    # 1) Go to ManageDatabases
    driver.get("https://pallprod.enlabel.com/Collaboration/ManageDatabases/ManageDatabases.aspx")
    _wait_ready_and_ajax(driver, 40)

    # 2) Find the tables grid (handles iframe if used)
    grid_tables_locator = (By.XPATH, "//*[contains(@id,'MainContent') and contains(@id,'gridTables')]")
    if not _switch_into_frame_if_needed(driver, grid_tables_locator, probe_timeout=3):
        raise TimeoutException("Could not locate the 'gridTables' on ManageDatabases page.")

    # 3) Click the link in the second data row (fallback to first visible row if needed)
    row2_link_xpath = "//*[@id[contains(.,'gridTables')]]//*[contains(@id,'__1')]//td[1]//a"
    link_elems = driver.find_elements(By.XPATH, row2_link_xpath)
    if not link_elems:
        row2_link_xpath = "//*[@id[contains(.,'gridTables')]]//tr[contains(@class,'rgRow') or contains(@class,'rgAltRow')][1]//td[1]//a"
        link_elems = driver.find_elements(By.XPATH, row2_link_xpath)
    if not link_elems:
        raise TimeoutException("No row link found in gridTables.")
    driver.execute_script("arguments[0].click();", link_elems[0])

    # 4) Wait for the records view to load (grid or filters)
    _wait_ready_and_ajax(driver, 40)
    driver.switch_to.default_content()

    # Accept any of these as landing targets
    targets = [
        (By.XPATH, "//*[contains(@id,'gridDbRecords')]"),
        (By.XPATH, "//*[contains(@id,'FilterControl_ddlOperand1')]"),
        (By.XPATH, "//*[contains(@id,'FilterControl_ddlColumn1')]"),
        (By.XPATH, "//*[contains(@class,'rgCommandRow') or contains(@class,'rgCommandCell') or contains(@id,'Command')]"),
    ]
    # Switch into the context that has any target
    found_context = False
    for loc in targets:
        if _switch_into_frame_if_needed(driver, loc, probe_timeout=2):
            found_context = True
            break
    if not found_context:
        raise TimeoutException("After opening the record, no grid/filters/command area detected.")

    # 5) If a command area exists and needs a click (optional), you can keep or remove this block.
    #    It is safe to skip if filters are already visible on your tenant.
    try:
        cmd = WebDriverWait(driver, 4).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(@id,'gridCommand') or contains(@class,'rgCommandRow') or contains(@class,'rgCommandCell') or contains(@id,'Command')]")
            )
        )
        anchors = [a for a in cmd.find_elements(By.XPATH, ".//a[normalize-space()]") if a.is_displayed() and a.is_enabled()]
        if len(anchors) >= 2:
            try:
                anchors[1].click()
            except Exception:
                driver.execute_script("arguments[0].click();", anchors[1])
            _wait_ready_and_ajax(driver, 15)
    except TimeoutException:
        # No command bar; filters may already be visible.
        pass

    # 6) Set filter dropdowns
    operand_dd = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(@id,'FilterControl_ddlOperand1')]"))
    )
    column_dd = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(@id,'FilterControl_ddlColumn1')]"))
    )
    Select(operand_dd).select_by_index(1)  # 2nd option
    Select(column_dd).select_by_index(8)   # 9th option


    lot_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_MainContent_FilterControl_txtValue1']"))
    )

    lot_input.clear()  # optional, clears any pre-filled text
    lot_input.send_keys(LOT_NUMBER)

    find_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
        By.XPATH, "//*[@id='ctl00_MainContent_FilterControl_btnFind']"
    )))

    find_button.click()
    time.sleep(2)

    production_number_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_MainContent_gridDbRecords_ctl00__0']/td[2]/nobr"))
    )

    production_number = production_number_element.get_attribute("textContent").strip()

    print("Extracted text: ", production_number)

    time.sleep(20)
    print("Navigation and production number extraction test completed.")

def search_label(driver, wait):

    
    print("Starting searching for labels")

def main():
    options = Options()

    driver = webdriver.Edge(options=options)


    wait = WebDriverWait(driver, 20)  # explicit wait for up to 20 seconds

    # 1) Open login page
    driver.get(LOGIN_URL)
    time.sleep(1)
    # 2) Wait for username field and enter text
    username_field = wait.until(EC.presence_of_element_located(USERNAME_LOCATOR))
    username_field.clear()
    username_field.send_keys(USERNAME)
    time.sleep(2)
    # 3) Wait for password field and enter text
    password_field = wait.until(EC.presence_of_element_located(PASSWORD_LOCATOR))
    password_field.clear()
    password_field.send_keys(PASSWORD)
    time.sleep(2)
    # 4) Click the submit/login button
    submit_btn = wait.until(EC.element_to_be_clickable(SUBMIT_LOCATOR))
    submit_btn.click()
    # 5) Optional: confirm login succeeded (example: wait for a dashboard element)
    # Replace with a locator that only appears after successful login.
    # DASHBOARD_LOCATOR = (By.CSS_SELECTOR, "nav .dashboard")
    # wait.until(EC.presence_of_element_located(DASHBOARD_LOCATOR))
    # Simple confirmation by printing the page title:
    print("Login attempted. Current page title:", driver.title)
    time.sleep(1)
    navigate_after_login(driver, wait)


    time.sleep(12)

if __name__ == "__main__":
    main()
