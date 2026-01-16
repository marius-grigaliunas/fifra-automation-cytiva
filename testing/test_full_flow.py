"""
Full flow test script for label automation.
This test combines all individual test steps into one complete workflow:
1. Search for label using production number
2. Open label preview
3. Resize and zoom in
4. Verify label content (item number, lot number, EPA number)
5. Print to PDF if verification passes
6. Save PDF file with item_number_lot_number format

Usage:
    python test_full_flow.py <item_number> <lot_number> <production_number>
    
Example:
    python test_full_flow.py 7005235 900100796 900030718
"""

import time
import sys
import os
from pathlib import Path
from datetime import datetime
import ctypes
from ctypes import wintypes
import re

# Selenium imports for browser automation
from selenium import webdriver
from selenium.webdriver.ie.options import Options as IEOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

# GUI automation imports
try:
    from pywinauto import Desktop
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False
    print("ERROR: pywinauto not installed. Install with: pip install pywinauto")
    sys.exit(1)

try:
    import pyautogui
    PYAutoGUI_AVAILABLE = True
except ImportError:
    PYAutoGUI_AVAILABLE = False
    print("ERROR: pyautogui not installed. Install with: pip install pyautogui")
    sys.exit(1)

try:
    from PIL import Image
    import numpy as np
    PILLOW_AVAILABLE = True
    NUMPY_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False
    NUMPY_AVAILABLE = False
    print("ERROR: Pillow or numpy not installed. Install with: pip install Pillow numpy")
    sys.exit(1)

try:
    import cv2
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False
    print("WARNING: opencv-python not installed. Some features may not work optimally")

try:
    import pytesseract
    PYTESSERACT_AVAILABLE = True
except ImportError:
    PYTESSERACT_AVAILABLE = False
    print("ERROR: pytesseract not installed. Install with: pip install pytesseract")
    sys.exit(1)

# Configuration
LOGIN_URL = "https://pallprod.enlabel.com/Login.aspx?ReturnUrl=%2f"
USERNAME = "mariusg"
PASSWORD = "LietuvaTevyneMusu!123"
LABEL_SEARCH_URL = "https://pallprod.enlabel.com/ProductionPrint/PrintTypes/PrintByOrder/PrintStart.aspx?ServiceId=10"

# Locators
PRODUCTION_NUMBER_INPUT_ID = "ctl00_MainContent__txtORDER_NUMBER"
SEARCH_BUTTON_ID = "ctl00_MainContent__btnNext"
RESULTS_TABLE_ID = "ctl00_MainContent_gridLabels"

# Logging setup
LOG_DIR = Path(__file__).parent
LOG_FILE = LOG_DIR / "test_full_flow_log.txt"
PROJECT_ROOT = LOG_DIR.parent
CONFIG_DIR = PROJECT_ROOT / "config"

# Button coordinates (from test_preview_window.py)
ABSOLUTE_PRINT_COORDS = (493, 421)


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


# ============================================================================
# Selenium Helper Functions (from test_label_search.py)
# ============================================================================

def _wait_ready_and_ajax(driver, timeout=30):
    """Wait for document.readyState == 'complete' and for jQuery to be idle (if present)."""
    w = WebDriverWait(driver, timeout)
    w.until(lambda d: d.execute_script("return document.readyState") == "complete")
    try:
        w.until(lambda d: d.execute_script("return (window.jQuery ? jQuery.active : 0) === 0"))
    except Exception:
        pass  # jQuery not present on all pages


def _switch_into_frame_if_needed(driver, locator, probe_timeout=2):
    """Ensure Selenium is in the DOM context that contains `locator`."""
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
        
        # Ensure we're in default content (not in an iframe)
        driver.switch_to.default_content()
        time.sleep(0.5)
        
        log(log_file, "Entering username...")
        username_field = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtUserName")))
        
        # Click to focus the field first
        try:
            username_field.click()
        except Exception:
            driver.execute_script("arguments[0].focus();", username_field)
        
        time.sleep(0.3)
        username_field.clear()
        time.sleep(0.2)
        username_field.send_keys(USERNAME)
        
        # Verify username was entered
        entered_username = username_field.get_attribute("value")
        log(log_file, f"Username entered: '{USERNAME}' (verified: '{entered_username}')")
        if entered_username != USERNAME:
            log(log_file, f"WARNING: Username value mismatch. Expected: '{USERNAME}', Got: '{entered_username}'", "WARNING")
            # Try JavaScript as fallback
            driver.execute_script("arguments[0].value = arguments[1];", username_field, USERNAME)
            time.sleep(0.3)
            entered_username = username_field.get_attribute("value")
            log(log_file, f"After JavaScript set: '{entered_username}'")
        
        time.sleep(1)
        
        log(log_file, "Entering password...")
        password_field = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtPassword")))
        
        # Click to focus the field first
        try:
            password_field.click()
        except Exception:
            driver.execute_script("arguments[0].focus();", password_field)
        
        time.sleep(0.3)
        password_field.clear()
        time.sleep(0.2)
        password_field.send_keys(PASSWORD)
        
        # Verify password was entered (check if field is not empty)
        entered_password = password_field.get_attribute("value")
        if entered_password:
            log(log_file, f"Password entered: {len(entered_password)} characters")
        else:
            log(log_file, "WARNING: Password field appears empty", "WARNING")
            # Try JavaScript as fallback
            driver.execute_script("arguments[0].value = arguments[1];", password_field, PASSWORD)
            time.sleep(0.3)
            entered_password = password_field.get_attribute("value")
            log(log_file, f"After JavaScript set: {len(entered_password) if entered_password else 0} characters")
        
        time.sleep(1)
        
        log(log_file, "Clicking login button...")
        submit_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnLogin")))
        try:
            submit_btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", submit_btn)
        
        _wait_ready_and_ajax(driver, 30)
        time.sleep(2)
        
        log(log_file, f"Login completed. Current page: {driver.title}")
        log(log_file, f"Current URL: {driver.current_url}")
        return True
    except Exception as e:
        log(log_file, f"Login failed: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
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
        
        input_locator = (By.ID, PRODUCTION_NUMBER_INPUT_ID)
        if _switch_into_frame_if_needed(driver, input_locator, probe_timeout=5):
            log(log_file, "Production number input field found")
        else:
            log(log_file, "WARNING: Production number input field not found", "WARNING")
        
        return True
    except Exception as e:
        log(log_file, f"Navigation failed: {type(e).__name__}: {str(e)}", "ERROR")
        return False


def enter_production_number(driver, wait, production_number, log_file):
    """Enter production number in the search field."""
    log(log_file, "\n" + "="*60)
    log(log_file, f"STEP 3: Entering production number: {production_number}")
    log(log_file, "="*60)
    
    try:
        driver.switch_to.default_content()
        input_locator = (By.ID, PRODUCTION_NUMBER_INPUT_ID)
        
        if not _switch_into_frame_if_needed(driver, input_locator, probe_timeout=5):
            log(log_file, "ERROR: Could not locate production number input field", "ERROR")
            return False
        
        prod_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(input_locator)
        )
        
        prod_input.clear()
        time.sleep(0.5)
        prod_input.send_keys(production_number)
        time.sleep(1)
        
        entered_value = prod_input.get_attribute("value")
        log(log_file, f"Value in input field: '{entered_value}'")
        
        if entered_value == production_number:
            log(log_file, "✓ Production number entered successfully")
            return True
        else:
            log(log_file, f"WARNING: Entered value does not match expected", "WARNING")
            return False
            
    except Exception as e:
        log(log_file, f"Error entering production number: {type(e).__name__}: {str(e)}", "ERROR")
        return False


def click_search_button(driver, wait, log_file):
    """Click the search/next button."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 4: Clicking search/next button")
    log(log_file, "="*60)
    
    try:
        log(log_file, f"Looking for search button (ID: {SEARCH_BUTTON_ID})...")
        search_button = None
        
        try:
            search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, SEARCH_BUTTON_ID))
            )
            log(log_file, f"Found search button by ID: {SEARCH_BUTTON_ID}")
        except TimeoutException:
            log(log_file, f"Search button not found by ID, trying alternative methods...", "WARNING")
            buttons = driver.find_elements(By.TAG_NAME, "input")
            for btn in buttons:
                btn_id = btn.get_attribute("id") or ""
                btn_type = btn.get_attribute("type") or ""
                btn_value = btn.get_attribute("value") or ""
                if btn_type in ["submit", "button"] and any(word in btn_value.lower() for word in ["next", "search", "find", "go"]):
                    search_button = btn
                    break
        
        if not search_button:
            log(log_file, "ERROR: Could not locate search button", "ERROR")
            return False
        
        log(log_file, "Clicking search button...")
        try:
            search_button.click()
        except Exception:
            driver.execute_script("arguments[0].click();", search_button)
        
        log(log_file, "Waiting for results to load...")
        time.sleep(3)
        _wait_ready_and_ajax(driver, 30)
        log(log_file, "✓ Search button clicked")
        return True
        
    except Exception as e:
        log(log_file, f"Error clicking search button: {type(e).__name__}: {str(e)}", "ERROR")
        return False


def extract_result_rows(driver, wait, log_file):
    """Extract all result rows from the label search results table."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 5: Extracting result rows")
    log(log_file, "="*60)
    
    try:
        driver.switch_to.default_content()
        result_rows = []
        row_index = 0
        
        while row_index <= 20:  # Safety limit
            patterns = [
                f"//*[@id='{RESULTS_TABLE_ID}_ctl00__{row_index}']",
                f"//tr[contains(@id,'gridLabels') and contains(@id,'__{row_index}')]",
            ]
            
            found = False
            for pattern in patterns:
                try:
                    row = driver.find_element(By.XPATH, pattern)
                    if row.is_displayed():
                        result_rows.append(row)
                        found = True
                        log(log_file, f"Found result row {row_index}")
                        break
                except NoSuchElementException:
                    continue
            
            if not found:
                break
            
            row_index += 1
        
        log(log_file, f"Found {len(result_rows)} result row(s)")
        return result_rows
        
    except Exception as e:
        log(log_file, f"Error extracting result rows: {type(e).__name__}: {str(e)}", "ERROR")
        return []


def click_preview_button(driver, wait, result_index, log_file):
    """Click preview button for a specific result row."""
    log(log_file, "\n" + "="*60)
    log(log_file, f"STEP 6: Clicking preview button for result row {result_index}")
    log(log_file, "="*60)
    
    try:
        ctl_index = 4 + (result_index * 2)
        preview_button_id = f"ctl00_MainContent_gridLabels_ctl00_ctl{ctl_index:02d}_btnPreview"
        log(log_file, f"Looking for preview button (ID: {preview_button_id})...")
        
        preview_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, preview_button_id))
        )
        log(log_file, f"Found preview button: {preview_button_id}")
        
        if not preview_button.is_displayed():
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", preview_button)
            time.sleep(1)
        
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, preview_button_id))
        )
        
        log(log_file, "Clicking preview button...")
        preview_button.click()
        
        log(log_file, "Preview button clicked, waiting for preview window to open...")
        time.sleep(5)  # Wait for preview window to open
        
        log(log_file, "✓ Preview button click completed")
        return True
        
    except Exception as e:
        log(log_file, f"Error clicking preview button: {type(e).__name__}: {str(e)}", "ERROR")
        return False


# ============================================================================
# Preview Window Functions (from test_preview_window.py)
# ============================================================================

def find_preview_window(log_file, timeout=30):
    """Find the preview window using pywinauto."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 7: Finding preview window")
    log(log_file, "="*60)
    
    title_patterns = [
        "enLabel Global Services - Work - Microsoft Edge",
        "enLabel Global Services - Internet Explorer",
        "enLabel Global Services",
    ]
    
    class_patterns = ["IEFrame", "Chrome_WidgetWin_1"]
    excluded_title_keywords = ["cursor", "vscode", "visual studio", "notepad++", "sublime", "fifra-automation"]
    
    start_time = time.time()
    desktop = Desktop(backend="win32")
    
    while time.time() - start_time < timeout:
        try:
            windows = desktop.windows()
            
            for window in windows:
                try:
                    if not window.is_visible():
                        continue
                    
                    title = window.window_text()
                    class_name = window.class_name()
                    title_lower = title.lower()
                    
                    if any(keyword in title_lower for keyword in excluded_title_keywords):
                        continue
                    
                    for pattern in class_patterns:
                        if pattern.lower() in class_name.lower():
                            if "chrome_widgetwin" in class_name.lower():
                                has_edge = "microsoft edge" in title_lower or ("edge" in title_lower and "microsoft" in title_lower)
                                if "enlabel" in title_lower and has_edge:
                                    log(log_file, f"✓ Found preview window:")
                                    log(log_file, f"    Title: '{title}'")
                                    log(log_file, f"    Class: '{class_name}'")
                                    return window
                    
                    for pattern in title_patterns:
                        if pattern.lower() in title_lower:
                            log(log_file, f"✓ Found preview window by title pattern:")
                            log(log_file, f"    Title: '{title}'")
                            log(log_file, f"    Class: '{class_name}'")
                            return window
                            
                except Exception:
                    continue
            
            time.sleep(0.2)
            
        except Exception as e:
            log(log_file, f"Error during search: {e}", "WARNING")
            time.sleep(0.2)
    
    log(log_file, f"Preview window not found after {timeout} seconds", "ERROR")
    return None


def resize_preview_window(log_file, start_x=850, start_y=827, end_x=1450, end_y=868):
    """Resize the preview window by dragging the bottom-right corner."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 8: Resizing preview window")
    log(log_file, "="*60)
    
    try:
        log(log_file, f"Dragging window from ({start_x}, {start_y}) to ({end_x}, {end_y})")
        pyautogui.moveTo(start_x, start_y)
        time.sleep(0.1)
        
        drag_x = end_x - start_x
        drag_y = end_y - start_y
        pyautogui.drag(drag_x, drag_y, duration=0.5, button='left')
        
        log(log_file, "✓ Window resize completed")
        time.sleep(0.5)
        return True
        
    except Exception as e:
        log(log_file, f"Error resizing window: {type(e).__name__}: {str(e)}", "ERROR")
        return False


def click_zoom_in_button(log_file, x=540, y=420, times=3):
    """Click the zoom in button multiple times."""
    log(log_file, "\n" + "="*60)
    log(log_file, f"STEP 9: Clicking zoom in button {times} times")
    log(log_file, "="*60)
    
    try:
        for i in range(times):
            pyautogui.click(x, y)
            log(log_file, f"  Click {i+1}/{times} at ({x}, {y})")
            time.sleep(0.2)
        
        log(log_file, f"✓ Clicked zoom in button {times} times")
        time.sleep(0.3)
        return True
        
    except Exception as e:
        log(log_file, f"Error clicking zoom in button: {type(e).__name__}: {str(e)}", "ERROR")
        return False


def click_print_button(window, log_file):
    """Click the print button."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 12: Clicking print button")
    log(log_file, "="*60)
    
    try:
        window_rect = window.rectangle()
        absolute_x, absolute_y = ABSOLUTE_PRINT_COORDS
        
        log(log_file, f"Using absolute screen coordinates: ({absolute_x}, {absolute_y})")
        
        try:
            window.set_focus()
            time.sleep(0.1)
        except Exception as e:
            log(log_file, f"Could not focus window: {e}", "WARNING")
        
        pyautogui.click(absolute_x, absolute_y)
        log(log_file, f"✓ Clicked print button at ({absolute_x}, {absolute_y})")
        
        time.sleep(1)  # Wait for printer dialog to open
        return True
        
    except Exception as e:
        log(log_file, f"Error clicking print button: {type(e).__name__}: {str(e)}", "ERROR")
        return False


def handle_printer_dialog(log_file):
    """Handle the Windows printer selection dialog."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 13: Handling printer selection dialog")
    log(log_file, "="*60)
    
    try:
        time.sleep(1)
        
        # Try to find and focus printer dialog
        if PYWINAUTO_AVAILABLE:
            try:
                desktop = Desktop(backend="win32")
                printer_dialogs = [
                    w for w in desktop.windows()
                    if w.is_visible() and (
                        "print" in w.window_text().lower() or
                        "printer" in w.window_text().lower() or
                        w.class_name() == "#32770"
                    )
                ]
                
                if printer_dialogs:
                    printer_dialog = printer_dialogs[0]
                    log(log_file, f"Found printer dialog: '{printer_dialog.window_text()}'")
                    printer_dialog.set_focus()
                    time.sleep(0.2)
            except Exception as e:
                log(log_file, f"Could not find printer dialog window: {e}", "WARNING")
        
        log(log_file, "Typing 'Microsoft Print to PDF' to select printer...")
        pyautogui.typewrite("Microsoft Print to PDF", interval=0.05)
        time.sleep(0.5)
        
        pyautogui.press('enter')
        time.sleep(0.3)
        
        log(log_file, "Pressing Alt+O to click OK button...")
        pyautogui.hotkey('alt', 'o')
        time.sleep(0.5)
        
        log(log_file, "Pressing Enter to confirm...")
        pyautogui.press('enter')
        
        time.sleep(1)
        log(log_file, "✓ Printer selection confirmed")
        return True
        
    except Exception as e:
        log(log_file, f"Error handling printer dialog: {type(e).__name__}: {str(e)}", "ERROR")
        return False


# ============================================================================
# Label Verification Functions (from test_label_verification.py)
# ============================================================================

def find_tesseract_path(log_file):
    """Find Tesseract executable path on Windows."""
    if PYTESSERACT_AVAILABLE:
        try:
            current_cmd = pytesseract.pytesseract.tesseract_cmd
            if current_cmd and os.path.exists(current_cmd):
                return current_cmd
        except:
            pass
    
    common_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
        str(CONFIG_DIR / "tesseract.exe"),
    ]
    
    for path in common_paths:
        if os.path.exists(path):
            log(log_file, f"Found Tesseract at: {path}")
            return path
    
    return None


def configure_tesseract(log_file):
    """Configure pytesseract to use Tesseract executable."""
    if not PYTESSERACT_AVAILABLE:
        return False
    
    try:
        tesseract_path = find_tesseract_path(log_file)
        if tesseract_path:
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
            log(log_file, f"Configured pytesseract to use: {tesseract_path}")
            return True
        else:
            log(log_file, "ERROR: Could not find Tesseract executable", "ERROR")
            return False
    except Exception as e:
        log(log_file, f"Error configuring Tesseract: {type(e).__name__}: {str(e)}", "ERROR")
        return False


def normalize_ocr_text(text):
    """Normalize common OCR errors in text."""
    if not text:
        return text
    
    text = text.replace('—', '-').replace('–', '-').replace('―', '-')
    
    def fix_s_in_numbers(match):
        prefix = match.group(1) or ''
        digits = match.group(2)
        suffix = match.group(3) or ''
        if suffix.startswith('m') and len(suffix) > 1:
            return prefix + '6' + digits + '0' + suffix[1:]
        return prefix + '6' + digits + suffix
    
    text = re.sub(r'(\b|[A-Za-z])([sS])(\d{1,2}([mM][0-9A-Za-z]|))', fix_s_in_numbers, text)
    text = re.sub(r'\b([sS])(\d)', r'6\2', text)
    text = re.sub(r'(\d)([mM])(\d)', r'\g<1>0\3', text)
    text = re.sub(r'(\d)([lL])(\d)', r'\g<1>1\3', text)
    text = re.sub(r'(\d)([O])(\d)', r'\g<1>0\3', text)
    text = re.sub(r'(\d)([I])(\d)', r'\g<1>1\3', text)
    
    return text


def extract_text_with_ocr(window, log_file):
    """Extract text from the preview window using OCR."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 10: Extracting text using OCR")
    log(log_file, "="*60)
    
    try:
        log(log_file, "Capturing screenshot of preview window...")
        window_rect = window.rectangle()
        
        screenshot_pil = pyautogui.screenshot(region=(
            window_rect.left,
            window_rect.top,
            window_rect.width(),
            window_rect.height()
        ))
        
        screenshot_path = LOG_DIR / f"label_verification_screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        screenshot_pil.save(screenshot_path)
        log(log_file, f"Screenshot saved to: {screenshot_path}")
        
        image_for_ocr = screenshot_pil
        if CV2_AVAILABLE and NUMPY_AVAILABLE:
            try:
                log(log_file, "Preprocessing image for OCR...")
                img_array = np.array(screenshot_pil)
                img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
                gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
                _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                image_for_ocr = Image.fromarray(thresh)
            except Exception as e:
                log(log_file, f"Image preprocessing failed: {e}", "WARNING")
        
        log(log_file, "Extracting text using OCR...")
        ocr_config = r'--oem 3 --psm 6'
        extracted_text = pytesseract.image_to_string(image_for_ocr, config=ocr_config)
        
        if extracted_text and len(extracted_text.strip()) > 0:
            normalized_text = normalize_ocr_text(extracted_text)
            if normalized_text != extracted_text:
                log(log_file, f"Applied normalization (fixed common OCR errors)")
                extracted_text = normalized_text
            
            log(log_file, f"✓ Successfully extracted text using OCR ({len(extracted_text)} characters)")
            log(log_file, f"Extracted text preview: '{extracted_text[:200]}...'")
            return extracted_text
        else:
            log(log_file, "OCR returned empty text, trying alternative config...", "WARNING")
            ocr_config_alt = r'--oem 3 --psm 11'
            extracted_text = pytesseract.image_to_string(image_for_ocr, config=ocr_config_alt)
            
            if extracted_text and len(extracted_text.strip()) > 0:
                normalized_text = normalize_ocr_text(extracted_text)
                extracted_text = normalized_text
                log(log_file, f"✓ Successfully extracted text using alternative config")
                return extracted_text
            else:
                log(log_file, "OCR returned empty text", "ERROR")
                return ""
                
    except Exception as e:
        log(log_file, f"Error during OCR extraction: {type(e).__name__}: {str(e)}", "ERROR")
        return ""


def verify_item_number(text, item_number, log_file):
    """Verify that the extracted text contains the item number."""
    log(log_file, "\n" + "="*60)
    log(log_file, "VERIFICATION: Item Number")
    log(log_file, "="*60)
    log(log_file, f"Expected item number: '{item_number}'")
    
    if not text:
        log(log_file, "No text available for verification", "ERROR")
        return False
    
    text_lower = text.lower()
    item_number_lower = item_number.lower()
    
    if item_number_lower in text_lower:
        log(log_file, f"✓ Item number found in text")
        return True
    else:
        log(log_file, f"✗ Item number '{item_number}' not found", "ERROR")
        return False


def verify_lot_number(text, lot_number, log_file):
    """Verify that the extracted text contains the lot number."""
    log(log_file, "\n" + "="*60)
    log(log_file, "VERIFICATION: Lot Number")
    log(log_file, "="*60)
    log(log_file, f"Expected lot number: '{lot_number}'")
    
    if not text:
        log(log_file, "No text available for verification", "ERROR")
        return False
    
    text_lower = text.lower()
    lot_number_lower = lot_number.lower()
    
    if lot_number_lower in text_lower:
        log(log_file, f"✓ Lot number found in text")
        return True
    else:
        log(log_file, f"✗ Lot number '{lot_number}' not found", "ERROR")
        return False


def verify_epa_number(text, log_file):
    """Verify that the extracted text contains 'EPA'."""
    log(log_file, "\n" + "="*60)
    log(log_file, "VERIFICATION: EPA Number")
    log(log_file, "="*60)
    
    if not text:
        log(log_file, "No text available for verification", "ERROR")
        return False
    
    text_lower = text.lower()
    
    if "epa" in text_lower:
        log(log_file, f"✓ 'EPA' found in text")
        return True
    else:
        log(log_file, f"✗ 'EPA' not found", "ERROR")
        return False


# ============================================================================
# File Save Dialog Functions (from test_file_save_dialog.py)
# ============================================================================

def find_save_dialog(log_file, timeout=30):
    """Find the Windows file save dialog using pywinauto."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 14: Finding file save dialog")
    log(log_file, "="*60)
    
    dialog_titles = [
        "Save As",
        "Save File",
        "Save Print Output As",
    ]
    
    dialog_classes = ["#32770"]
    excluded_keywords = ["cursor", "vscode", "visual studio", "notepad++", "sublime", "fifra-automation"]
    
    start_time = time.time()
    desktop = Desktop(backend="win32")
    
    while time.time() - start_time < timeout:
        try:
            windows = desktop.windows()
            
            for window in windows:
                try:
                    if not window.is_visible():
                        continue
                    
                    title = window.window_text()
                    class_name = window.class_name()
                    title_lower = title.lower()
                    
                    if any(keyword in title_lower for keyword in excluded_keywords):
                        continue
                    
                    if "#32770" in class_name or class_name == "#32770":
                        try:
                            descendants = window.descendants()
                            has_filename_edit = any("Edit" in str(ctrl.control_type()) for ctrl in descendants)
                            has_save_button = any("Save" in str(ctrl.window_text()) and "Button" in str(ctrl.control_type()) for ctrl in descendants)
                            
                            if has_filename_edit or has_save_button:
                                log(log_file, f"✓ Found save dialog:")
                                log(log_file, f"    Title: '{title}'")
                                log(log_file, f"    Class: '{class_name}'")
                                return window
                        except Exception:
                            pass
                    
                    for dialog_title in dialog_titles:
                        if title == dialog_title or title.lower() == dialog_title.lower():
                            log(log_file, f"✓ Found save dialog by title:")
                            log(log_file, f"    Title: '{title}'")
                            log(log_file, f"    Class: '{class_name}'")
                            return window
                            
                except Exception:
                    continue
            
            time.sleep(0.5)
            
        except Exception as e:
            log(log_file, f"Error during search: {e}", "WARNING")
            time.sleep(0.5)
    
    log(log_file, f"Save dialog not found after {timeout} seconds", "ERROR")
    return None


def handle_save_dialog(dialog, save_path, log_file):
    """Handle the file save dialog using pywinauto."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 15: Handling save dialog")
    log(log_file, "="*60)
    
    try:
        dialog.set_focus()
        time.sleep(0.5)
        
        # Set directory path via Alt+D
        log(log_file, "Setting directory path using Alt+D...")
        pyautogui.hotkey('alt', 'd')
        time.sleep(0.5)
        
        directory_path = str(save_path.parent.absolute())
        log(log_file, f"Typing directory path: {directory_path}")
        pyautogui.typewrite(directory_path, interval=0.05)
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(1.5)
        
        # Set filename
        pyautogui.press('tab')
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.2)
        
        filename = save_path.name
        log(log_file, f"Typing filename: {filename}")
        pyautogui.typewrite(filename, interval=0.05)
        time.sleep(0.5)
        
        # Save
        pyautogui.press('enter')
        log(log_file, "✓ Pressed Enter to save file")
        time.sleep(2)
        
        if save_path.exists():
            log(log_file, f"✓ File saved successfully: {save_path}")
            return True
        else:
            log(log_file, f"WARNING: File not found after save: {save_path}", "WARNING")
            return False
        
    except Exception as e:
        log(log_file, f"Error handling save dialog: {type(e).__name__}: {str(e)}", "ERROR")
        return False


# ============================================================================
# Main Function
# ============================================================================

def main():
    """Main test function."""
    log_file = setup_logging()
    
    log(log_file, "="*60)
    log(log_file, "Full Flow Label Automation Test")
    log(log_file, "="*60)
    log(log_file, f"Test started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(log_file, "")
    
    # Check command line arguments
    if len(sys.argv) < 4:
        log(log_file, "ERROR: Missing required arguments", "ERROR")
        log(log_file, "Usage: python test_full_flow.py <item_number> <lot_number> <production_number>", "ERROR")
        log(log_file, "Example: python test_full_flow.py 7005235 900100796 900030718", "ERROR")
        log_file.close()
        sys.exit(1)
    
    item_number = sys.argv[1]
    lot_number = sys.argv[2]
    production_number = sys.argv[3]
    
    log(log_file, f"Item number: '{item_number}'")
    log(log_file, f"Lot number: '{lot_number}'")
    log(log_file, f"Production number: '{production_number}'")
    log(log_file, "")
    
    # Configure Tesseract
    if not configure_tesseract(log_file):
        log(log_file, "Could not configure Tesseract, aborting", "ERROR")
        log_file.close()
        sys.exit(1)
    
    driver = None
    result = False
    
    try:
        # Setup browser
        log(log_file, "Setting up Edge browser in Internet Explorer mode...")
        options = IEOptions()
        options.attach_to_edge_chrome = True
        options.ignore_protected_mode_settings = True
        options.ignore_zoom_level = True
        options.require_window_focus = False
        options.ensure_clean_session = True
        
        driver = webdriver.Ie(options=options)
        wait = WebDriverWait(driver, 20)
        log(log_file, "Browser started successfully")
        log(log_file, "")
        
        # Step 1: Login
        if not login(driver, wait, log_file):
            log(log_file, "Login failed, aborting", "ERROR")
            return
        
        # Step 2: Navigate to label search
        if not navigate_to_label_search(driver, wait, log_file):
            log(log_file, "Navigation failed, aborting", "ERROR")
            return
        
        # Step 3: Enter production number
        if not enter_production_number(driver, wait, production_number, log_file):
            log(log_file, "Failed to enter production number, aborting", "ERROR")
            return
        
        # Step 4: Click search button
        if not click_search_button(driver, wait, log_file):
            log(log_file, "Failed to click search button, aborting", "ERROR")
            return
        
        # Step 5: Extract result rows
        result_rows = extract_result_rows(driver, wait, log_file)
        if not result_rows:
            log(log_file, "No result rows found, aborting", "ERROR")
            return
        
        # Step 6: Click preview button for first result
        if not click_preview_button(driver, wait, 0, log_file):
            log(log_file, "Failed to click preview button, aborting", "ERROR")
            return
        
        # Step 7: Find preview window
        preview_window = find_preview_window(log_file, timeout=30)
        if not preview_window:
            log(log_file, "Could not find preview window, aborting", "ERROR")
            return
        
        # Step 8: Resize preview window
        resize_preview_window(log_file, start_x=850, start_y=827, end_x=1450, end_y=868)
        
        # Step 9: Click zoom in button
        click_zoom_in_button(log_file, x=540, y=420, times=3)
        
        # Step 10: Extract text using OCR
        extracted_text = extract_text_with_ocr(preview_window, log_file)
        if not extracted_text:
            log(log_file, "Could not extract text from label, aborting", "ERROR")
            return
        
        # Step 11: Verify label content
        log(log_file, "\n" + "="*60)
        log(log_file, "STEP 11: Verifying label content")
        log(log_file, "="*60)
        
        item_verified = verify_item_number(extracted_text, item_number, log_file)
        lot_verified = verify_lot_number(extracted_text, lot_number, log_file)
        epa_verified = verify_epa_number(extracted_text, log_file)
        
        log(log_file, "\n" + "="*60)
        log(log_file, "VERIFICATION SUMMARY")
        log(log_file, "="*60)
        log(log_file, f"Item Number: {'✓ PASS' if item_verified else '✗ FAIL'}")
        log(log_file, f"Lot Number: {'✓ PASS' if lot_verified else '✗ FAIL'}")
        log(log_file, f"EPA Number: {'✓ PASS' if epa_verified else '✗ FAIL'}")
        log(log_file, "")
        
        if not (item_verified and lot_verified and epa_verified):
            log(log_file, "✗ VERIFICATION FAILED - Label does not match expected values", "ERROR")
            log(log_file, "Aborting print operation", "ERROR")
            result = False
        else:
            log(log_file, "✓ ALL VERIFICATIONS PASSED - Proceeding with print", "INFO")
            
            # Step 12: Click print button
            if not click_print_button(preview_window, log_file):
                log(log_file, "Failed to click print button", "ERROR")
                result = False
            else:
                # Step 13: Handle printer dialog
                if not handle_printer_dialog(log_file):
                    log(log_file, "Failed to handle printer dialog", "ERROR")
                    result = False
                else:
                    # Step 14: Find save dialog
                    save_dialog = find_save_dialog(log_file, timeout=30)
                    if not save_dialog:
                        log(log_file, "Could not find save dialog", "ERROR")
                        result = False
                    else:
                        # Step 15: Handle save dialog
                        # Create filename: item_number_lot_number.pdf
                        filename = f"{item_number}_{lot_number}.pdf"
                        save_path = LOG_DIR / filename
                        
                        log(log_file, f"Target filename: {filename}")
                        log(log_file, f"Target path: {save_path}")
                        
                        if handle_save_dialog(save_dialog, save_path, log_file):
                            log(log_file, "✓ File saved successfully")
                            result = True
                        else:
                            log(log_file, "Failed to save file", "ERROR")
                            result = False
        
        log(log_file, "")
        log(log_file, "="*60)
        if result:
            log(log_file, "✓ TEST COMPLETED SUCCESSFULLY")
        else:
            log(log_file, "✗ TEST FAILED")
        log(log_file, "="*60)
        
    except Exception as e:
        log(log_file, f"Test failed with error: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        result = False
    finally:
        if driver:
            log(log_file, "Closing browser...")
            try:
                driver.quit()
            except:
                pass
        log_file.close()
        print(f"\nTest log saved to: {LOG_FILE}")
    
    sys.exit(0 if result else 1)


if __name__ == "__main__":
    main()