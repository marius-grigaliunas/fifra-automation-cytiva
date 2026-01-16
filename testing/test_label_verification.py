"""
Test script for label verification system.
Tests if a label in an opened preview window contains the required information:
- Item number (must match provided item number)
- Lot number (must match provided lot number)
- EPA number (must contain "EPA")

Usage:
    python test_label_verification.py <item_number> <lot_number>
    
Example:
    python test_label_verification.py 6401-1167T 900114574
"""

import time
import sys
import os
from pathlib import Path
from datetime import datetime

try:
    from pywinauto import Desktop
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False
    print("WARNING: pywinauto not installed. Install with: pip install pywinauto")

try:
    import pyautogui
    PYAutoGUI_AVAILABLE = True
except ImportError:
    PYAutoGUI_AVAILABLE = False
    print("WARNING: pyautogui not installed. Install with: pip install pyautogui")

try:
    from PIL import Image
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False
    print("WARNING: Pillow not installed. Install with: pip install Pillow")

try:
    import cv2
    import numpy as np
    CV2_AVAILABLE = True
    NUMPY_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False
    NUMPY_AVAILABLE = False
    print("WARNING: opencv-python or numpy not installed. Install with: pip install opencv-python numpy")

try:
    import pytesseract
    PYTESSERACT_AVAILABLE = True
except ImportError:
    PYTESSERACT_AVAILABLE = False
    print("WARNING: pytesseract not installed. Install with: pip install pytesseract")
    print("NOTE: Also requires Tesseract OCR binary installed on system")
    print("     Download from: https://github.com/UB-Mannheim/tesseract/wiki")

# Logging setup
LOG_DIR = Path(__file__).parent
LOG_FILE = LOG_DIR / "test_label_verification_log.txt"
PROJECT_ROOT = LOG_DIR.parent
CONFIG_DIR = PROJECT_ROOT / "config"


def find_tesseract_path(log_file):
    """
    Find Tesseract executable path on Windows without requiring PATH.
    Checks common installation locations and config file.
    
    Returns:
        str: Path to tesseract.exe or None if not found
    """
    tesseract_exe = "tesseract.exe"
    possible_paths = []
    
    # Check if already configured
    if PYTESSERACT_AVAILABLE:
        try:
            current_cmd = pytesseract.pytesseract.tesseract_cmd
            if current_cmd and os.path.exists(current_cmd):
                log(log_file, f"Tesseract already configured: {current_cmd}")
                return current_cmd
        except:
            pass
    
    # Common Windows installation paths
    common_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
        os.path.expanduser(r"~\AppData\Local\Tesseract-OCR\tesseract.exe"),
        r"C:\Tesseract-OCR\tesseract.exe",
        r"D:\Program Files\Tesseract-OCR\tesseract.exe",
        r"D:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    
    # Check for tesseract in project config directory
    config_tesseract = CONFIG_DIR / "tesseract.exe"
    if config_tesseract.exists():
        common_paths.insert(0, str(config_tesseract))
    
    # Check if there's a config file specifying the path
    config_file = CONFIG_DIR / "tesseract_path.txt"
    if config_file.exists():
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                custom_path = f.read().strip()
                if custom_path and os.path.exists(custom_path):
                    common_paths.insert(0, custom_path)
        except Exception as e:
            log(log_file, f"Error reading tesseract_path.txt: {e}", "WARNING")
    
    # Check environment variable
    env_path = os.environ.get('TESSERACT_CMD')
    if env_path and os.path.exists(env_path):
        common_paths.insert(0, env_path)
    
    # Search for tesseract in common paths
    for path in common_paths:
        if os.path.exists(path):
            log(log_file, f"Found Tesseract at: {path}")
            return path
    
    # Last resort: try to find tesseract in PATH (might work if admin installed it)
    import shutil
    path_tesseract = shutil.which(tesseract_exe)
    if path_tesseract:
        log(log_file, f"Found Tesseract in PATH: {path_tesseract}")
        return path_tesseract
    
    log(log_file, "Tesseract not found in common locations", "WARNING")
    log(log_file, "To configure manually:", "WARNING")
    log(log_file, "  1. Create config/tesseract_path.txt with the full path to tesseract.exe", "WARNING")
    log(log_file, "  2. Set TESSERACT_CMD environment variable to the tesseract.exe path", "WARNING")
    log(log_file, "  3. Place tesseract.exe in config/ folder", "WARNING")
    return None


def configure_tesseract(log_file):
    """
    Configure pytesseract to use Tesseract executable found on system.
    """
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
            log(log_file, "Please specify Tesseract path using one of these methods:", "ERROR")
            log(log_file, "  1. Create config/tesseract_path.txt with full path to tesseract.exe", "ERROR")
            log(log_file, "  2. Set TESSERACT_CMD environment variable", "ERROR")
            log(log_file, "  3. Place tesseract.exe in config/ folder", "ERROR")
            return False
    except Exception as e:
        log(log_file, f"Error configuring Tesseract: {type(e).__name__}: {str(e)}", "ERROR")
        return False


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


def find_preview_window(log_file, timeout=30):
    """
    Find the preview window using pywinauto.
    
    Returns:
        Window object if found, None otherwise
    """
    if not PYWINAUTO_AVAILABLE:
        log(log_file, "pywinauto not available", "ERROR")
        return None
    
    title_patterns = [
        "enLabel Global Services - Work - Microsoft Edge",
        "enLabel Global Services - Internet Explorer",
        "enLabel Global Services",
        "Label Preview",
        "Print Preview",
        "enLabel",
    ]
    
    class_patterns = [
        "IEFrame",
        "Chrome_WidgetWin_1",
    ]
    
    excluded_title_keywords = ["cursor", "vscode", "visual studio", "notepad++", "sublime", "fifra-automation"]
    
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 1: Finding preview window")
    log(log_file, "="*60)
    
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
                    
                    # Skip excluded windows
                    if any(keyword in title_lower for keyword in excluded_title_keywords):
                        continue
                    
                    # Check class patterns first
                    for pattern in class_patterns:
                        if pattern.lower() in class_name.lower():
                            if "chrome_widgetwin" in class_name.lower():
                                has_edge = "microsoft edge" in title_lower or ("edge" in title_lower and "microsoft" in title_lower)
                                if "enlabel" in title_lower and has_edge:
                                    log(log_file, f"✓ Found window by class pattern '{pattern}':")
                                    log(log_file, f"    Title: '{title}'")
                                    log(log_file, f"    Class: '{class_name}'")
                                    return window
                            else:
                                if "enlabel" in title_lower or "internet explorer" in title_lower or "preview" in title_lower:
                                    log(log_file, f"✓ Found window by class pattern '{pattern}':")
                                    log(log_file, f"    Title: '{title}'")
                                    log(log_file, f"    Class: '{class_name}'")
                                    return window
                    
                    # Check title patterns
                    for pattern in title_patterns:
                        if pattern.lower() in title_lower:
                            log(log_file, f"✓ Found window by title pattern '{pattern}':")
                            log(log_file, f"    Title: '{title}'")
                            log(log_file, f"    Class: '{class_name}'")
                            return window
                            
                except Exception:
                    continue
            
            time.sleep(0.5)
            
        except Exception as e:
            log(log_file, f"Error during search: {e}", "WARNING")
            time.sleep(0.5)
    
    log(log_file, f"Preview window not found after {timeout} seconds", "ERROR")
    return None


def extract_text_with_ocr(window, log_file):
    """
    Extract text from the preview window using OCR.
    Uses OpenCV for image preprocessing and pytesseract for OCR.
    The label content is rendered as an image in an ActiveX control, so OCR is required.
    
    Args:
        window: Window object from pywinauto
        log_file: Log file handle
    
    Returns:
        str: Extracted text from OCR or empty string if extraction fails
    """
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 2: Extracting text using OCR (screenshot + pytesseract)")
    log(log_file, "="*60)
    
    if not PYTESSERACT_AVAILABLE:
        log(log_file, "ERROR: pytesseract is not installed", "ERROR")
        log(log_file, "Install with: pip install pytesseract", "ERROR")
        log(log_file, "Also requires Tesseract OCR binary:", "ERROR")
        log(log_file, "  Download from: https://github.com/UB-Mannheim/tesseract/wiki", "ERROR")
        return ""
    
    if not PILLOW_AVAILABLE or not PYAutoGUI_AVAILABLE:
        log(log_file, "ERROR: Pillow or pyautogui not available for screenshot", "ERROR")
        return ""
    
    try:
        # Step 1: Capture screenshot of the window
        log(log_file, "Capturing screenshot of preview window...")
        window_rect = window.rectangle()
        
        screenshot_pil = pyautogui.screenshot(region=(
            window_rect.left,
            window_rect.top,
            window_rect.width(),
            window_rect.height()
        ))
        
        # Save screenshot for debugging
        screenshot_path = LOG_DIR / f"label_verification_screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        screenshot_pil.save(screenshot_path)
        log(log_file, f"Screenshot saved to: {screenshot_path}")
        log(log_file, f"Screenshot size: {screenshot_pil.width}x{screenshot_pil.height} pixels")
        
        # Step 2: Preprocess image with OpenCV for better OCR results (optimized for speed)
        image_for_ocr = screenshot_pil
        if CV2_AVAILABLE and NUMPY_AVAILABLE:
            try:
                log(log_file, "Preprocessing image for OCR (optimized single pass)...")
                # Convert PIL to OpenCV format
                img_array = np.array(screenshot_pil)
                img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
                
                # Convert to grayscale
                gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
                
                # Use Otsu's thresholding (fast and effective)
                _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                
                # Convert back to PIL Image for pytesseract
                image_for_ocr = Image.fromarray(thresh)
                
                # Save processed image for debugging (optional, can be removed for speed)
                processed_path = LOG_DIR / f"label_verification_processed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                image_for_ocr.save(processed_path)
                log(log_file, f"Processed image saved: {processed_path}")
            except Exception as e:
                log(log_file, f"Image preprocessing failed, using original image: {e}", "WARNING")
        
        # Step 3: Extract text using pytesseract (single optimized pass)
        log(log_file, "Extracting text using OCR...")
        try:
            # Use PSM 6 (uniform block) - works well for labels, fast
            ocr_config = r'--oem 3 --psm 6'
            extracted_text = pytesseract.image_to_string(image_for_ocr, config=ocr_config)
            
            if extracted_text and len(extracted_text.strip()) > 0:
                log(log_file, f"✓ Successfully extracted text using OCR ({len(extracted_text)} characters)")
                
                # Normalize OCR text to fix common errors
                normalized_text = normalize_ocr_text(extracted_text)
                if normalized_text != extracted_text:
                    log(log_file, f"Applied normalization (fixed common OCR errors)")
                    extracted_text = normalized_text
                
                log(log_file, f"Extracted text preview: '{extracted_text[:300]}...' (truncated)")
                log(log_file, f"Full extracted text:")
                log(log_file, "-" * 60)
                log(log_file, extracted_text)
                log(log_file, "-" * 60)
                
                return extracted_text
            else:
                # Fallback: try PSM 11 (sparse text) if first attempt failed
                log(log_file, "First OCR attempt returned empty text, trying alternative config...")
                ocr_config_alt = r'--oem 3 --psm 11'
                extracted_text = pytesseract.image_to_string(image_for_ocr, config=ocr_config_alt)
                
                if extracted_text and len(extracted_text.strip()) > 0:
                    log(log_file, f"✓ Successfully extracted text using alternative config ({len(extracted_text)} characters)")
                    
                    # Normalize OCR text
                    normalized_text = normalize_ocr_text(extracted_text)
                    if normalized_text != extracted_text:
                        log(log_file, f"Applied normalization (fixed common OCR errors)")
                        extracted_text = normalized_text
                    
                    log(log_file, f"Extracted text: '{extracted_text[:300]}...' (truncated)")
                    return extracted_text
                else:
                    log(log_file, "OCR returned empty text", "WARNING")
                    log(log_file, "This may indicate:")
                    log(log_file, "  1. Label content is not visible in the screenshot")
                    log(log_file, "  2. Image quality is too low")
                    log(log_file, "  3. Text font/style is not recognized by OCR")
                    return ""
                
        except pytesseract.TesseractNotFoundError:
            log(log_file, "ERROR: Tesseract OCR binary not found", "ERROR")
            log(log_file, "Please install Tesseract OCR:", "ERROR")
            log(log_file, "  Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki", "ERROR")
            log(log_file, "  Make sure Tesseract is in your PATH or configure pytesseract.pytesseract.tesseract_cmd", "ERROR")
            return ""
        except Exception as e:
            log(log_file, f"Error during OCR extraction: {type(e).__name__}: {str(e)}", "ERROR")
            import traceback
            log(log_file, traceback.format_exc(), "ERROR")
            return ""
        
    except Exception as e:
        log(log_file, f"Error during OCR text extraction: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return ""




def normalize_ocr_text(text):
    """
    Normalize common OCR errors in text.
    Fixes common character misreadings like 6→s, 0→O, -→—, etc.
    
    Args:
        text: Text extracted from OCR
    
    Returns:
        str: Normalized text
    """
    if not text:
        return text
    
    import re
    
    # Replace em dash and en dash with regular hyphen
    text = text.replace('—', '-').replace('–', '-').replace('―', '-')
    
    # Common OCR character confusions in item numbers and codes
    # Pattern-based fixes for common OCR errors
    
    # Fix: 's' before digits in number sequences (6→s error)
    # Pattern: s followed by 1-2 digits, possibly followed by more alphanumeric
    # Example: s4m→640, s11→611
    def fix_s_in_numbers(match):
        prefix = match.group(1) or ''
        digits = match.group(2)
        suffix = match.group(3) or ''
        # If suffix is 'm' followed by number-like, might be '0'
        if suffix.startswith('m') and len(suffix) > 1:
            # Could be s4m → 640, check if what follows 'm' looks numeric
            return prefix + '6' + digits + '0' + suffix[1:]
        return prefix + '6' + digits + suffix
    
    # Fix patterns like "s4", "s4m", "s11" in alphanumeric sequences
    text = re.sub(r'(\b|[A-Za-z])([sS])(\d{1,2}([mM][0-9A-Za-z]|))', fix_s_in_numbers, text)
    
    # Fix standalone 's' before digits (if previous fix didn't catch it)
    text = re.sub(r'\b([sS])(\d)', r'6\2', text)
    
    # Fix: lowercase 'm' in middle of numbers might be '0'
    # Pattern: digit + m + digit (like 4m1 → 401)
    # Use \g<1> syntax to avoid \10 being interpreted as group 10
    text = re.sub(r'(\d)([mM])(\d)', r'\g<1>0\3', text)
    
    # Fix: 'l' (lowercase L) → '1' in number contexts
    # Pattern: digit + l + digit (like 6l1 → 611)
    text = re.sub(r'(\d)([lL])(\d)', r'\g<1>1\3', text)
    
    # Fix: 'O' (capital O) → '0' in pure number sequences
    # Pattern: surrounded by digits (like 6O1 → 601)
    text = re.sub(r'(\d)([O])(\d)', r'\g<1>0\3', text)
    
    # Fix: 'I' (capital i) → '1' in number sequences
    text = re.sub(r'(\d)([I])(\d)', r'\g<1>1\3', text)
    
    return text


def verify_item_number(text, item_number, log_file):
    """
    Verify that the extracted text contains the item number.
    
    Args:
        text: Extracted text from label
        item_number: Expected item number
        log_file: Log file handle
    
    Returns:
        bool: True if item number is found, False otherwise
    """
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 3: Verifying Item Number")
    log(log_file, "="*60)
    log(log_file, f"Expected item number: '{item_number}'")
    
    if not text:
        log(log_file, "No text available for verification", "ERROR")
        return False
    
    # Search for item number (case-insensitive)
    text_lower = text.lower()
    item_number_lower = item_number.lower()
    
    if item_number_lower in text_lower:
        # Find the context around the item number
        index = text_lower.find(item_number_lower)
        start = max(0, index - 30)
        end = min(len(text), index + len(item_number) + 30)
        context = text[start:end]
        log(log_file, f"✓ Item number found in text")
        log(log_file, f"  Context: '...{context}...'")
        return True
    else:
        log(log_file, f"✗ Item number '{item_number}' not found in extracted text", "ERROR")
        log(log_file, f"  Extracted text: '{text}'")
        return False


def verify_lot_number(text, lot_number, log_file):
    """
    Verify that the extracted text contains the lot number.
    
    Args:
        text: Extracted text from label
        lot_number: Expected lot number
        log_file: Log file handle
    
    Returns:
        bool: True if lot number is found, False otherwise
    """
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 4: Verifying Lot Number")
    log(log_file, "="*60)
    log(log_file, f"Expected lot number: '{lot_number}'")
    
    if not text:
        log(log_file, "No text available for verification", "ERROR")
        return False
    
    # Search for lot number (case-insensitive)
    text_lower = text.lower()
    lot_number_lower = lot_number.lower()
    
    if lot_number_lower in text_lower:
        # Find the context around the lot number
        index = text_lower.find(lot_number_lower)
        start = max(0, index - 30)
        end = min(len(text), index + len(lot_number) + 30)
        context = text[start:end]
        log(log_file, f"✓ Lot number found in text")
        log(log_file, f"  Context: '...{context}...'")
        return True
    else:
        log(log_file, f"✗ Lot number '{lot_number}' not found in extracted text", "ERROR")
        log(log_file, f"  Extracted text: '{text}'")
        return False


def verify_epa_number(text, log_file):
    """
    Verify that the extracted text contains "EPA" (case-insensitive).
    
    Args:
        text: Extracted text from label
        log_file: Log file handle
    
    Returns:
        bool: True if "EPA" is found, False otherwise
    """
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 5: Verifying EPA Number")
    log(log_file, "="*60)
    
    if not text:
        log(log_file, "No text available for verification", "ERROR")
        return False
    
    # Search for "EPA" (case-insensitive)
    text_lower = text.lower()
    
    if "epa" in text_lower:
        # Find the context around "EPA"
        index = text_lower.find("epa")
        start = max(0, index - 30)
        end = min(len(text), index + 40)
        context = text[start:end]
        log(log_file, f"✓ 'EPA' found in text")
        log(log_file, f"  Context: '...{context}...'")
        return True
    else:
        log(log_file, f"✗ 'EPA' not found in extracted text", "ERROR")
        log(log_file, f"  Extracted text: '{text}'")
        return False


def main():
    """Main test function."""
    log_file = setup_logging()
    
    log(log_file, "="*60)
    log(log_file, "Label Verification Test Script")
    log(log_file, "="*60)
    log(log_file, f"Test started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(log_file, "")
    
    # Check dependencies
    if not PYWINAUTO_AVAILABLE:
        log(log_file, "ERROR: pywinauto is not installed", "ERROR")
        log_file.close()
        return
    
    if not PYTESSERACT_AVAILABLE:
        log(log_file, "ERROR: pytesseract is not installed", "ERROR")
        log(log_file, "Install with: pip install pytesseract", "ERROR")
        log(log_file, "Also requires Tesseract OCR binary:", "ERROR")
        log(log_file, "  Download from: https://github.com/UB-Mannheim/tesseract/wiki", "ERROR")
        log_file.close()
        return
    
    # Configure Tesseract path (required if not in PATH)
    log(log_file, "Configuring Tesseract path...")
    if not configure_tesseract(log_file):
        log(log_file, "Could not configure Tesseract path", "ERROR")
        log_file.close()
        return
    
    if not PILLOW_AVAILABLE or not PYAutoGUI_AVAILABLE:
        log(log_file, "ERROR: Pillow or pyautogui is not installed", "ERROR")
        log(log_file, "Install with: pip install Pillow pyautogui", "ERROR")
        log_file.close()
        return
    
    # Get item number and lot number from command line arguments
    if len(sys.argv) < 3:
        log(log_file, "ERROR: Missing required arguments", "ERROR")
        log(log_file, "Usage: python test_label_verification.py <item_number> <lot_number>", "ERROR")
        log(log_file, "Example: python test_label_verification.py 6401-1167T 900114574", "ERROR")
        log_file.close()
        return
    
    item_number = sys.argv[1]
    lot_number = sys.argv[2]
    
    log(log_file, f"Item number to verify: '{item_number}'")
    log(log_file, f"Lot number to verify: '{lot_number}'")
    log(log_file, "")
    
    # Ask user to confirm preview window is open
    log(log_file, "IMPORTANT: This test assumes a preview window is already open.")
    log(log_file, "Please ensure a label preview window is open before continuing.")
    log(log_file, "")
    response = input("Is the preview window open? (y/n): ").strip().lower()
    
    if response != 'y':
        log(log_file, "Test aborted - preview window not confirmed as open", "ERROR")
        log_file.close()
        return
    
    log(log_file, "User confirmed preview window is open")
    log(log_file, "")
    
    try:
        # Step 1: Find preview window
        preview_window = find_preview_window(log_file, timeout=10)
        
        if not preview_window:
            log(log_file, "Could not find preview window", "ERROR")
            log(log_file, "Please ensure a preview window is open and try again", "ERROR")
            log_file.close()
            return
        
        # Step 2: Extract text from preview window using OCR
        extracted_text = extract_text_with_ocr(preview_window, log_file)
        
        if not extracted_text:
            log(log_file, "ERROR: Could not extract text from preview window using OCR", "ERROR")
            log(log_file, "This may indicate:", "ERROR")
            log(log_file, "  1. Tesseract OCR binary is not installed or not in PATH", "ERROR")
            log(log_file, "  2. Label content is not visible in the screenshot", "ERROR")
            log(log_file, "  3. OCR quality settings need adjustment", "ERROR")
            log(log_file, "")
            log(log_file, "Check the saved screenshot to verify label visibility", "ERROR")
            log_file.close()
            sys.exit(1)
        
        # Step 3-5: Verify label components
        item_verified = verify_item_number(extracted_text, item_number, log_file)
        lot_verified = verify_lot_number(extracted_text, lot_number, log_file)
        epa_verified = verify_epa_number(extracted_text, log_file)
        
        # Summary
        log(log_file, "\n" + "="*60)
        log(log_file, "VERIFICATION SUMMARY")
        log(log_file, "="*60)
        log(log_file, f"Item Number ('{item_number}'): {'✓ PASS' if item_verified else '✗ FAIL'}")
        log(log_file, f"Lot Number ('{lot_number}'): {'✓ PASS' if lot_verified else '✗ FAIL'}")
        log(log_file, f"EPA Number (contains 'EPA'): {'✓ PASS' if epa_verified else '✗ FAIL'}")
        log(log_file, "")
        
        if item_verified and lot_verified and epa_verified:
            log(log_file, "✓ ALL VERIFICATIONS PASSED - Label is valid")
            result = True
        else:
            log(log_file, "✗ VERIFICATION FAILED - Label is missing required information", "ERROR")
            result = False
        
        log(log_file, "="*60)
        
    except Exception as e:
        log(log_file, f"Test failed with error: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        result = False
    finally:
        log_file.close()
        print(f"\nTest log saved to: {LOG_FILE}")
    
    # Exit with appropriate code
    sys.exit(0 if result else 1)


if __name__ == "__main__":
    main()
