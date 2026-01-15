"""
Test script to verify Selenium can open Edge in Internet Explorer mode.
This tests the feasibility of the hybrid automation approach.

This script tests:
1. Opening Edge in IE mode
2. Logging into enLabel
3. Navigating to production search
4. Searching for a production number by lot number

Note: This requires IEDriverServer.exe to be installed and in PATH.
Download from: https://github.com/SeleniumHQ/selenium/releases
Look for "IEDriverServer" in the assets.

Usage:
    python test_edge_ie_mode.py [lot_number]
    
Example:
    python test_edge_ie_mode.py UE4376
"""

import time
import sys
import os
from selenium import webdriver
from selenium.webdriver.ie.options import Options as IEOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException


# Configuration - Update these or set via environment variables
LOGIN_URL = "https://pallprod.enlabel.com/Login.aspx?ReturnUrl=%2f"
USERNAME = os.getenv("ENLABEL_USERNAME", "mariusg")
PASSWORD = os.getenv("ENLABEL_PASSWORD", "LietuvaTevyneMusu!123")

# Locators
USERNAME_LOCATOR = (By.ID, "ctl00_ContentPlaceHolder1_txtUserName")
PASSWORD_LOCATOR = (By.ID, "ctl00_ContentPlaceHolder1_txtPassword")
SUBMIT_LOCATOR = (By.ID, "ctl00_ContentPlaceHolder1_btnLogin")


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


def login(driver, wait):
    """Perform login to enLabel."""
    print("\n" + "="*60)
    print("STEP 1: Logging in to enLabel")
    print("="*60)
    
    driver.get(LOGIN_URL)
    _wait_ready_and_ajax(driver, 30)
    
    print("Entering username...")
    username_field = wait.until(EC.presence_of_element_located(USERNAME_LOCATOR))
    username_field.clear()
    username_field.send_keys(USERNAME)
    time.sleep(1)
    
    print("Entering password...")
    password_field = wait.until(EC.presence_of_element_located(PASSWORD_LOCATOR))
    password_field.clear()
    password_field.send_keys(PASSWORD)
    time.sleep(1)
    
    print("Clicking login button...")
    submit_btn = wait.until(EC.element_to_be_clickable(SUBMIT_LOCATOR))
    submit_btn.click()
    
    # Wait for login to complete
    _wait_ready_and_ajax(driver, 30)
    time.sleep(2)
    
    print(f"Login completed. Current page: {driver.title}")
    print(f"Current URL: {driver.current_url}")


def navigate_to_production_search(driver, wait):
    """Navigate to ManageDatabases and open the database for searching."""
    print("\n" + "="*60)
    print("STEP 2: Navigating to production search")
    print("="*60)
    
    # Navigate to ManageDatabases
    manage_db_url = "https://pallprod.enlabel.com/Collaboration/ManageDatabases/ManageDatabases.aspx"
    print(f"Navigating to: {manage_db_url}")
    driver.get(manage_db_url)
    _wait_ready_and_ajax(driver, 40)
    
    # Find the tables grid (handles iframe if used)
    grid_tables_locator = (By.XPATH, "//*[contains(@id,'MainContent') and contains(@id,'gridTables')]")
    if not _switch_into_frame_if_needed(driver, grid_tables_locator, probe_timeout=3):
        raise TimeoutException("Could not locate the 'gridTables' on ManageDatabases page.")
    
    # Click the link in the second data row (fallback to first visible row if needed)
    print("Clicking on database entry...")
    row2_link_xpath = "//*[@id[contains(.,'gridTables')]]//*[contains(@id,'__1')]//td[1]//a"
    link_elems = driver.find_elements(By.XPATH, row2_link_xpath)
    if not link_elems:
        row2_link_xpath = "//*[@id[contains(.,'gridTables')]]//tr[contains(@class,'rgRow') or contains(@class,'rgAltRow')][1]//td[1]//a"
        link_elems = driver.find_elements(By.XPATH, row2_link_xpath)
    if not link_elems:
        raise TimeoutException("No row link found in gridTables.")
    
    driver.execute_script("arguments[0].click();", link_elems[0])
    
    # Wait for the records view to load
    _wait_ready_and_ajax(driver, 40)
    driver.switch_to.default_content()
    
    # Find the filter context
    targets = [
        (By.XPATH, "//*[contains(@id,'gridDbRecords')]"),
        (By.XPATH, "//*[contains(@id,'FilterControl_ddlOperand1')]"),
        (By.XPATH, "//*[contains(@id,'FilterControl_ddlColumn1')]"),
        (By.XPATH, "//*[contains(@class,'rgCommandRow') or contains(@class,'rgCommandCell') or contains(@id,'Command')]"),
    ]
    
    found_context = False
    for loc in targets:
        if _switch_into_frame_if_needed(driver, loc, probe_timeout=2):
            found_context = True
            break
    if not found_context:
        raise TimeoutException("After opening the record, no grid/filters/command area detected.")
    
    # Click command area if needed to show filters
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
        pass  # Filters may already be visible
    
    print("Navigation to production search completed.")


def search_production_by_lot(driver, wait, lot_number):
    """Search for production number using lot number."""
    print("\n" + "="*60)
    print(f"STEP 3: Searching for production number with lot: {lot_number}")
    print("="*60)
    
    # Set filter dropdowns
    print("Setting filter dropdowns...")
    operand_dd = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(@id,'FilterControl_ddlOperand1')]"))
    )
    column_dd = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(@id,'FilterControl_ddlColumn1')]"))
    )
    
    Select(operand_dd).select_by_index(1)  # 2nd option
    Select(column_dd).select_by_index(8)   # 9th option
    
    # Enter lot number
    print(f"Entering lot number: {lot_number}")
    lot_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_MainContent_FilterControl_txtValue1']"))
    )
    lot_input.clear()
    lot_input.send_keys(lot_number)
    time.sleep(1)
    
    # Click find button
    print("Clicking find button...")
    find_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((
        By.XPATH, "//*[@id='ctl00_MainContent_FilterControl_btnFind']"
    )))
    find_button.click()
    
    # Wait for results
    time.sleep(3)
    _wait_ready_and_ajax(driver, 30)
    
    # Extract production number
    print("Extracting production number from results...")
    try:
        production_number_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_MainContent_gridDbRecords_ctl00__0']/td[2]/nobr"))
        )
        production_number = production_number_element.get_attribute("textContent").strip()
        print(f"\n✓ SUCCESS: Found production number: {production_number}")
        return production_number
    except TimeoutException:
        print("\n✗ ERROR: Could not find production number in results.")
        print("The search may have returned no results, or the element locator may need adjustment.")
        return None


def test_edge_ie_mode_with_login_and_search(lot_number):
    """
    Test opening Edge in IE mode, logging in, and searching for production number.
    
    This uses InternetExplorerDriver with attachToEdgeChrome() to run Edge in IE mode.
    """
    print("="*60)
    print("Edge IE Mode Test - Login and Production Search")
    print("="*60)
    print("Setting up Edge with Internet Explorer mode...")
    print("Note: This requires IEDriverServer.exe to be in your PATH")
    
    # Configure Internet Explorer options to attach to Edge
    options = IEOptions()
    
    # This is the key method: attach to Edge Chrome instead of IE
    options.attach_to_edge_chrome = True
    
    # Additional IE options that may help
    options.ignore_protected_mode_settings = True
    options.ignore_zoom_level = True
    options.require_window_focus = False
    
    print("Starting Edge browser in IE mode...")
    print("Note: The browser window should open and you should see IE mode indicators.")
    
    driver = None
    try:
        # Create the driver using InternetExplorerDriver
        driver = webdriver.Ie(options=options)
        wait = WebDriverWait(driver, 20)
        
        # Perform login
        login(driver, wait)
        
        # Navigate to production search
        navigate_to_production_search(driver, wait)
        
        # Search for production number
        production_number = search_production_by_lot(driver, wait, lot_number)
        
        # Keep browser open for verification
        print("\n" + "="*60)
        print("Test completed. Browser will stay open for 30 seconds for verification...")
        print("="*60)
        time.sleep(30)
        
        return production_number
        
    except Exception as e:
        print(f"\n✗ ERROR: {type(e).__name__}: {str(e)}")
        print("\nTroubleshooting tips:")
        print("1. Ensure IEDriverServer.exe is installed and in PATH")
        print("   Download from: https://github.com/SeleniumHQ/selenium/releases")
        print("2. Ensure Edge browser is installed")
        print("3. Check if IE mode is enabled in Edge settings")
        print("4. Verify credentials are correct")
        print("5. Verify the lot number exists in the system")
        raise
    finally:
        if driver:
            print("\nClosing browser...")
            driver.quit()
            print("Browser closed successfully.")


if __name__ == "__main__":
    # Get lot number from command line argument or use default
    if len(sys.argv) > 1:
        lot_number = sys.argv[1]
    else:
        lot_number = input("Enter lot number to search for (or press Enter for default 'UE4376'): ").strip()
        if not lot_number:
            lot_number = "UE4376"
    
    print(f"\nUsing lot number: {lot_number}\n")
    
    try:
        production_number = test_edge_ie_mode_with_login_and_search(lot_number)
        if production_number:
            print(f"\n✓ Test completed successfully!")
            print(f"  Lot Number: {lot_number}")
            print(f"  Production Number: {production_number}")
        else:
            print(f"\n⚠ Test completed but production number was not found.")
    except Exception as e:
        print(f"\n✗ Test failed: {e}")
        sys.exit(1)
