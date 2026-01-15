"""
Test script for label search functionality.
Tests navigation to label search page, entering production number, clicking search,
extracting result rows, and clicking preview buttons.

Usage:
    python test_label_search.py [production_number]
    
Example:
    python test_label_search.py 900105880
"""

import time
import sys
import os
import config 
from pathlib import Path
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.ie.options import Options as IEOptions
from selenium.webdriver.ie.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, WebDriverException
from src.config_loader import get_config

# Configuration
LOGIN_URL = "https://pallprod.enlabel.com/Login.aspx?ReturnUrl=%2f"
USERNAME = ""
PASSWORD = ""
LABEL_SEARCH_URL = "https://pallprod.enlabel.com/ProductionPrint/PrintTypes/PrintByOrder/PrintStart.aspx?ServiceId=10"

# Locators (to be verified during testing)
PRODUCTION_NUMBER_INPUT_ID = "ctl00_MainContent__txtORDER_NUMBER"
SEARCH_BUTTON_ID = "ctl00_MainContent__btnNext"  # Needs verification
CLEAR_BUTTON_ID = "ctl00_MainContent__btnClear"  # Needs verification
RESULTS_TABLE_ID = "ctl00_MainContent_gridLabels"

# Logging setup
LOG_DIR = Path(__file__).parent
LOG_FILE = LOG_DIR / "test_label_search_log.txt"


def setup_logging():
    """Setup logging to file."""
    log_file = open(LOG_FILE, 'w', encoding='utf-8')
    return log_file


def log(log_file, message, level="INFO"):
    """Write log message to file and console."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] [{level}] {message}\n"
    log_file.write(log_entry)
    log_file.flush()
    print(f"[{level}] {message}")


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


def login(driver, wait, log_file):
    """Perform login to enLabel."""
    log(log_file, "="*60)
    log(log_file, "STEP 1: Logging in to enLabel")
    log(log_file, "="*60)
    
    try:
        log(log_file, f"Navigating to login URL: {LOGIN_URL}")
        driver.get(LOGIN_URL)
        _wait_ready_and_ajax(driver, 30)
        log(log_file, "Login page loaded successfully")
        
        log(log_file, "Entering username...")
        username_field = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtUserName")))
        username_field.clear()
        username_field.send_keys(USERNAME)
        log(log_file, f"Username entered: {USERNAME}")
        time.sleep(1)
        
        log(log_file, "Entering password...")
        password_field = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtPassword")))
        password_field.clear()
        password_field.send_keys(PASSWORD)
        log(log_file, "Password entered")
        time.sleep(1)
        
        log(log_file, "Clicking login button...")
        submit_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnLogin")))
        submit_btn.click()
        
        _wait_ready_and_ajax(driver, 30)
        time.sleep(2)
        
        log(log_file, f"Login completed. Current page: {driver.title}")
        log(log_file, f"Current URL: {driver.current_url}")
        return True
    except Exception as e:
        log(log_file, f"Login failed: {type(e).__name__}: {str(e)}", "ERROR")
        return False


def navigate_to_label_search(driver, wait, log_file):
    """Navigate to label search page."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 2: Navigating to label search page")
    log(log_file, "="*60)
    
    try:
        log(log_file, f"Navigating to label search URL: {LABEL_SEARCH_URL}")
        driver.get(LABEL_SEARCH_URL)
        _wait_ready_and_ajax(driver, 40)
        log(log_file, "Label search page loaded")
        log(log_file, f"Page title: {driver.title}")
        log(log_file, f"Current URL: {driver.current_url}")
        
        # Check if we need to switch to iframe
        log(log_file, "Checking for iframes...")
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        log(log_file, f"Found {len(iframes)} iframe(s)")
        
        # Try to locate production number input field
        log(log_file, f"Looking for production number input field (ID: {PRODUCTION_NUMBER_INPUT_ID})...")
        input_locator = (By.ID, PRODUCTION_NUMBER_INPUT_ID)
        
        if _switch_into_frame_if_needed(driver, input_locator, probe_timeout=5):
            log(log_file, "Production number input field found (switched to iframe if needed)")
        else:
            log(log_file, "WARNING: Production number input field not found with expected ID", "WARNING")
            log(log_file, "Attempting to find input field by alternative methods...")
            
            # Try to find any input field that might be the production number field
            try:
                inputs = driver.find_elements(By.TAG_NAME, "input")
                log(log_file, f"Found {len(inputs)} input field(s) on page")
                for i, inp in enumerate(inputs):
                    inp_id = inp.get_attribute("id") or "no-id"
                    inp_type = inp.get_attribute("type") or "no-type"
                    inp_name = inp.get_attribute("name") or "no-name"
                    log(log_file, f"  Input {i+1}: id='{inp_id}', type='{inp_type}', name='{inp_name}'")
            except Exception as e:
                log(log_file, f"Error finding input fields: {e}", "ERROR")
        
        return True
    except Exception as e:
        log(log_file, f"Navigation to label search failed: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def enter_production_number(driver, wait, production_number, log_file):
    """Enter production number in the search field."""
    log(log_file, "\n" + "="*60)
    log(log_file, f"STEP 3: Entering production number: {production_number}")
    log(log_file, "="*60)
    
    try:
        # Switch to correct context
        driver.switch_to.default_content()
        input_locator = (By.ID, PRODUCTION_NUMBER_INPUT_ID)
        
        if not _switch_into_frame_if_needed(driver, input_locator, probe_timeout=5):
            log(log_file, "ERROR: Could not locate production number input field", "ERROR")
            return False
        
        log(log_file, "Locating production number input field...")
        prod_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(input_locator)
        )
        
        log(log_file, "Clearing input field...")
        prod_input.clear()
        time.sleep(0.5)
        
        log(log_file, f"Entering production number: {production_number}")
        prod_input.send_keys(production_number)
        time.sleep(1)
        
        # Verify the value was entered
        entered_value = prod_input.get_attribute("value")
        log(log_file, f"Value in input field: '{entered_value}'")
        
        if entered_value == production_number:
            log(log_file, "✓ Production number entered successfully")
            return True
        else:
            log(log_file, f"WARNING: Entered value '{entered_value}' does not match expected '{production_number}'", "WARNING")
            return False
            
    except Exception as e:
        log(log_file, f"Error entering production number: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def click_search_button(driver, wait, log_file):
    """Click the search/next button."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 4: Clicking search/next button")
    log(log_file, "="*60)
    
    try:
        # Try to find search button by ID first
        log(log_file, f"Looking for search button (ID: {SEARCH_BUTTON_ID})...")
        search_button = None
        
        try:
            search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, SEARCH_BUTTON_ID))
            )
            log(log_file, f"Found search button by ID: {SEARCH_BUTTON_ID}")
        except TimeoutException:
            log(log_file, f"Search button not found by ID: {SEARCH_BUTTON_ID}", "WARNING")
            log(log_file, "Attempting to find search button by alternative methods...")
            
            # Try to find button by text or other attributes
            try:
                buttons = driver.find_elements(By.TAG_NAME, "input")
                log(log_file, f"Found {len(buttons)} input button(s)")
                for i, btn in enumerate(buttons):
                    btn_id = btn.get_attribute("id") or "no-id"
                    btn_type = btn.get_attribute("type") or "no-type"
                    btn_value = btn.get_attribute("value") or "no-value"
                    log(log_file, f"  Button {i+1}: id='{btn_id}', type='{btn_type}', value='{btn_value}'")
                    
                    # Look for buttons with "Next", "Search", "Find" in value
                    if btn_type == "submit" or btn_type == "button":
                        if any(word in btn_value.lower() for word in ["next", "search", "find", "go"]):
                            log(log_file, f"  → Potential search button found: id='{btn_id}', value='{btn_value}'")
                            search_button = btn
                            break
            except Exception as e:
                log(log_file, f"Error finding buttons: {e}", "ERROR")
        
        if not search_button:
            log(log_file, "ERROR: Could not locate search button", "ERROR")
            return False
        
        log(log_file, "Clicking search button...")
        try:
            search_button.click()
        except Exception:
            log(log_file, "Regular click failed, trying JavaScript click...")
            driver.execute_script("arguments[0].click();", search_button)
        
        log(log_file, "Waiting for results to load...")
        time.sleep(3)
        _wait_ready_and_ajax(driver, 30)
        log(log_file, "✓ Search button clicked, waiting for results")
        return True
        
    except Exception as e:
        log(log_file, f"Error clicking search button: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def extract_result_rows(driver, wait, log_file):
    """Extract all result rows from the label search results table."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 5: Extracting result rows")
    log(log_file, "="*60)
    
    try:
        # Switch to correct context
        driver.switch_to.default_content()
        
        # Look for results table
        log(log_file, f"Looking for results table (ID: {RESULTS_TABLE_ID})...")
        table_locator = (By.ID, RESULTS_TABLE_ID)
        
        if not _switch_into_frame_if_needed(driver, table_locator, probe_timeout=5):
            log(log_file, "Results table not found with expected ID", "WARNING")
            log(log_file, "Attempting to find table by alternative methods...")
            
            # Try to find any table that might contain results
            tables = driver.find_elements(By.TAG_NAME, "table")
            log(log_file, f"Found {len(tables)} table(s) on page")
            for i, table in enumerate(tables):
                table_id = table.get_attribute("id") or "no-id"
                table_class = table.get_attribute("class") or "no-class"
                log(log_file, f"  Table {i+1}: id='{table_id}', class='{table_class}'")
        
        # Try to find result rows
        log(log_file, "Looking for result rows...")
        
        # Pattern for result rows: gridLabels_ctl00__0, gridLabels_ctl00__1, etc.
        result_rows = []
        row_index = 0
        
        while True:
            # Try different XPath patterns
            patterns = [
                f"//*[@id='{RESULTS_TABLE_ID}_ctl00__{row_index}']",
                f"//tr[contains(@id,'gridLabels') and contains(@id,'__{row_index}')]",
                f"//*[contains(@id,'gridLabels') and contains(@id,'__{row_index}')]",
            ]
            
            found = False
            for pattern in patterns:
                try:
                    row = driver.find_element(By.XPATH, pattern)
                    if row.is_displayed():
                        result_rows.append(row)
                        found = True
                        log(log_file, f"Found result row {row_index} using pattern: {pattern}")
                        break
                except NoSuchElementException:
                    continue
            
            if not found:
                break
            
            row_index += 1
            if row_index > 20:  # Safety limit
                log(log_file, "Reached safety limit of 20 rows", "WARNING")
                break
        
        log(log_file, f"Found {len(result_rows)} result row(s)")
        
        # Extract details from each row
        for i, row in enumerate(result_rows):
            try:
                row_id = row.get_attribute("id") or "no-id"
                row_text = row.text[:100] if row.text else "no-text"  # First 100 chars
                log(log_file, f"  Row {i+1}: id='{row_id}', text preview='{row_text}...'")
            except Exception as e:
                log(log_file, f"  Row {i+1}: Error extracting details: {e}", "WARNING")
        
        return result_rows
        
    except Exception as e:
        log(log_file, f"Error extracting result rows: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return []


def click_preview_button(driver, wait, result_index, log_file):
    """Click preview button for a specific result row."""
    log(log_file, "\n" + "="*60)
    log(log_file, f"STEP 6: Clicking preview button for result row {result_index}")
    log(log_file, "="*60)
    
    try:
        # Preview button pattern: ctl00_MainContent_gridLabels_ctl00_ctl04_btnPreview (row 0)
        # Row 0: ctl04, Row 1: ctl06, Row 2: ctl08, etc.
        # Pattern: ctl04 + (row_index * 2) = ctl04, ctl06, ctl08, ctl10...
        ctl_index = 4 + (result_index * 2)
        
        preview_button_id = f"ctl00_MainContent_gridLabels_ctl00_ctl{ctl_index:02d}_btnPreview"
        log(log_file, f"Looking for preview button (ID: {preview_button_id})...")
        
        try:
            preview_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, preview_button_id))
            )
            log(log_file, f"Found preview button: {preview_button_id}")
            
            log(log_file, "Clicking preview button...")
            try:
                preview_button.click()
            except Exception:
                log(log_file, "Regular click failed, trying JavaScript click...")
                driver.execute_script("arguments[0].click();", preview_button)
            
            log(log_file, "Preview button clicked, waiting for preview window to open...")
            time.sleep(3)  # Give preview window time to open
            
            # Check if new window opened
            window_handles = driver.window_handles
            log(log_file, f"Current number of windows: {len(window_handles)}")
            for i, handle in enumerate(window_handles):
                driver.switch_to.window(handle)
                log(log_file, f"  Window {i+1}: title='{driver.title}', url='{driver.current_url}'")
            
            log(log_file, "✓ Preview button clicked successfully")
            return True
            
        except TimeoutException:
            log(log_file, f"Preview button not found with ID: {preview_button_id}", "WARNING")
            log(log_file, "Attempting to find preview button by alternative methods...")
            
            # Try to find button by text or XPath
            try:
                buttons = driver.find_elements(By.XPATH, f"//input[@type='button' or @type='submit']")
                log(log_file, f"Found {len(buttons)} button(s)")
                for i, btn in enumerate(buttons):
                    btn_id = btn.get_attribute("id") or "no-id"
                    btn_value = btn.get_attribute("value") or "no-value"
                    if "preview" in btn_id.lower() or "preview" in btn_value.lower():
                        log(log_file, f"  → Potential preview button: id='{btn_id}', value='{btn_value}'")
            except Exception as e:
                log(log_file, f"Error finding buttons: {e}", "ERROR")
            
            return False
            
    except Exception as e:
        log(log_file, f"Error clicking preview button: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def main():
    """Main test function."""
    log_file = setup_logging()
    
    log(log_file, "="*60)
    log(log_file, "Label Search Test Script")
    log(log_file, "="*60)
    log(log_file, f"Test started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(log_file, "")
    
    # Get production number from command line or use default
    if len(sys.argv) > 1:
        production_number = sys.argv[1]
    else:
        production_number = input("Enter production number to search for (or press Enter for default '900105880'): ").strip()
        if not production_number:
            production_number = "900105880"
    
    log(log_file, f"Using production number: {production_number}")
    log(log_file, "")
    
    # Setup Edge in IE mode
    log(log_file, "Setting up Edge browser in Internet Explorer mode...")
    options = IEOptions()
    options.attach_to_edge_chrome = True
    options.ignore_protected_mode_settings = True
    options.ignore_zoom_level = True
    options.require_window_focus = False
    options.ensure_clean_session = True
    
    driver = None
    try:
        log(log_file, "Starting Edge browser...")
        driver = webdriver.Ie(options=options)
        wait = WebDriverWait(driver, 20)
        log(log_file, "Browser started successfully")
        log(log_file, "")
        
        # Step 1: Login
        if not login(driver, wait, log_file):
            log(log_file, "Login failed, aborting test", "ERROR")
            return
        
        # Step 2: Navigate to label search
        if not navigate_to_label_search(driver, wait, log_file):
            log(log_file, "Navigation failed, aborting test", "ERROR")
            return
        
        # Step 3: Enter production number
        if not enter_production_number(driver, wait, production_number, log_file):
            log(log_file, "Failed to enter production number, aborting test", "ERROR")
            return
        
        # Step 4: Click search button
        if not click_search_button(driver, wait, log_file):
            log(log_file, "Failed to click search button, aborting test", "ERROR")
            return
        
        # Step 5: Extract result rows
        result_rows = extract_result_rows(driver, wait, log_file)
        if not result_rows:
            log(log_file, "No result rows found", "WARNING")
        else:
            log(log_file, f"Successfully extracted {len(result_rows)} result row(s)")
        
        # Step 6: Click preview button for first result (if available)
        if result_rows:
            log(log_file, "")
            log(log_file, "Attempting to click preview button for first result...")
            click_preview_button(driver, wait, 0, log_file)
            log(log_file, "")
            log(log_file, "Keeping browser open for 30 seconds for manual inspection...")
            time.sleep(30)
        else:
            log(log_file, "No results to preview")
            log(log_file, "Keeping browser open for 30 seconds for manual inspection...")
            time.sleep(30)
        
        log(log_file, "")
        log(log_file, "="*60)
        log(log_file, "Test completed successfully")
        log(log_file, "="*60)
        
    except Exception as e:
        log(log_file, f"Test failed with error: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
    finally:
        if driver:
            log(log_file, "Closing browser...")
            try:
                driver.quit()
            except:
                pass
        log_file.close()
        log_file = None
        print(f"\nTest log saved to: {LOG_FILE}")


if __name__ == "__main__":
    main()
