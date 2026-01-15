"""
Test script for preview window interaction using pywinauto.
Tests finding preview window, extracting text, clicking print button, and closing window.

Note: This script assumes a preview window is already open. You should run test_label_search.py
first to open a preview window, then run this script in a separate terminal.

Alternatively, you can manually open a preview window before running this script.

Usage:
    python test_preview_window.py
"""

import time
import sys
from pathlib import Path
from datetime import datetime

try:
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import ElementNotFoundError
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False
    print("WARNING: pywinauto not installed. Install with: pip install pywinauto")

# Logging setup
LOG_DIR = Path(__file__).parent
LOG_FILE = LOG_DIR / "test_preview_window_log.txt"


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


def list_all_windows(log_file):
    """List all open windows for debugging."""
    log(log_file, "\n" + "="*60)
    log(log_file, "Listing all open windows")
    log(log_file, "="*60)
    
    if not PYWINAUTO_AVAILABLE:
        log(log_file, "pywinauto not available, cannot list windows", "ERROR")
        return []
    
    try:
        desktop = Desktop(backend="win32")
        windows = desktop.windows()
        log(log_file, f"Found {len(windows)} window(s)")
        
        for i, window in enumerate(windows):
            try:
                title = window.window_text()
                class_name = window.class_name()
                is_visible = window.is_visible()
                log(log_file, f"  Window {i+1}:")
                log(log_file, f"    Title: '{title}'")
                log(log_file, f"    Class: '{class_name}'")
                log(log_file, f"    Visible: {is_visible}")
            except Exception as e:
                log(log_file, f"  Window {i+1}: Error getting info - {e}", "WARNING")
        
        return windows
    except Exception as e:
        log(log_file, f"Error listing windows: {type(e).__name__}: {str(e)}", "ERROR")
        return []


def find_preview_window(log_file, title_patterns=None, class_patterns=None, timeout=30):
    """
    Find the preview window using various methods.
    
    Args:
        log_file: Log file handle
        title_patterns: List of title patterns to search for
        class_patterns: List of class name patterns to search for
        timeout: Timeout in seconds
    
    Returns:
        Window object if found, None otherwise
    """
    if not PYWINAUTO_AVAILABLE:
        log(log_file, "pywinauto not available", "ERROR")
        return None
    
    if title_patterns is None:
        title_patterns = [
            "Label Preview",
            "Preview",
            "Label",
            "Print Preview",
            "enLabel",
        ]
    
    if class_patterns is None:
        class_patterns = [
            "Internet Explorer_Server",
            "IEFrame",
            "Shell Embedding",
            "ActiveX",
        ]
    
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 1: Finding preview window")
    log(log_file, "="*60)
    log(log_file, f"Searching with timeout: {timeout} seconds")
    log(log_file, f"Title patterns: {title_patterns}")
    log(log_file, f"Class patterns: {class_patterns}")
    
    start_time = time.time()
    desktop = Desktop(backend="win32")
    
    while time.time() - start_time < timeout:
        try:
            windows = desktop.windows()
            log(log_file, f"Checking {len(windows)} windows...")
            
            for window in windows:
                try:
                    if not window.is_visible():
                        continue
                    
                    title = window.window_text()
                    class_name = window.class_name()
                    
                    # Check title patterns
                    for pattern in title_patterns:
                        if pattern.lower() in title.lower():
                            log(log_file, f"✓ Found window by title pattern '{pattern}':")
                            log(log_file, f"    Title: '{title}'")
                            log(log_file, f"    Class: '{class_name}'")
                            return window
                    
                    # Check class patterns
                    for pattern in class_patterns:
                        if pattern.lower() in class_name.lower():
                            log(log_file, f"✓ Found window by class pattern '{pattern}':")
                            log(log_file, f"    Title: '{title}'")
                            log(log_file, f"    Class: '{class_name}'")
                            return window
                    
                except Exception as e:
                    continue
            
            time.sleep(1)
            
        except Exception as e:
            log(log_file, f"Error during search: {e}", "WARNING")
            time.sleep(1)
    
    log(log_file, f"Preview window not found after {timeout} seconds", "ERROR")
    log(log_file, "Available windows:")
    list_all_windows(log_file)
    return None


def extract_text_from_preview_window(window, log_file):
    """Extract text from preview window."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 2: Extracting text from preview window")
    log(log_file, "="*60)
    
    if not window:
        log(log_file, "No window provided", "ERROR")
        return None
    
    try:
        log(log_file, "Attempting to extract window text...")
        
        # Method 1: Get window text directly
        try:
            window_text = window.window_text()
            log(log_file, f"Window text (first 500 chars): {window_text[:500]}")
            if len(window_text) > 500:
                log(log_file, f"... (truncated, total length: {len(window_text)} chars)")
        except Exception as e:
            log(log_file, f"Could not get window text directly: {e}", "WARNING")
            window_text = ""
        
        # Method 2: Try to get text from child elements
        try:
            log(log_file, "Attempting to get text from child elements...")
            children = window.children()
            log(log_file, f"Found {len(children)} child element(s)")
            
            all_text = []
            for i, child in enumerate(children[:10]):  # Limit to first 10 children
                try:
                    child_text = child.window_text()
                    if child_text:
                        all_text.append(child_text)
                        log(log_file, f"  Child {i+1} text: {child_text[:100]}...")
                except:
                    pass
            
            if all_text:
                combined_text = " ".join(all_text)
                log(log_file, f"Combined child text (first 500 chars): {combined_text[:500]}")
                return combined_text
        except Exception as e:
            log(log_file, f"Could not get child text: {e}", "WARNING")
        
        # Method 3: Try to get element info
        try:
            log(log_file, "Attempting to get element info...")
            element_info = window.element_info
            log(log_file, f"Element name: {element_info.name}")
            log(log_file, f"Element class: {element_info.class_name}")
            log(log_file, f"Element control type: {element_info.control_type}")
        except Exception as e:
            log(log_file, f"Could not get element info: {e}", "WARNING")
        
        return window_text if window_text else None
        
    except Exception as e:
        log(log_file, f"Error extracting text: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return None


def find_print_button(window, log_file):
    """Find print button in preview window."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 3: Finding print button")
    log(log_file, "="*60)
    
    if not window:
        log(log_file, "No window provided", "ERROR")
        return None
    
    try:
        log(log_file, "Searching for print button...")
        
        # Method 1: Search by button text
        button_texts = ["Print", "Print...", "Print Label", "Print Preview"]
        for text in button_texts:
            try:
                button = window.child_window(title=text, control_type="Button")
                if button.exists():
                    log(log_file, f"✓ Found print button by text: '{text}'")
                    return button
            except:
                pass
        
        # Method 2: Search all buttons
        try:
            buttons = window.descendants(control_type="Button")
            log(log_file, f"Found {len(buttons)} button(s) in window")
            
            for i, button in enumerate(buttons):
                try:
                    button_text = button.window_text()
                    button_id = getattr(button, 'element_info', {}).get('name', 'no-name')
                    log(log_file, f"  Button {i+1}: text='{button_text}', id='{button_id}'")
                    
                    if "print" in button_text.lower():
                        log(log_file, f"  → Potential print button found: '{button_text}'")
                        return button
                except Exception as e:
                    log(log_file, f"  Button {i+1}: Error getting info - {e}", "WARNING")
        except Exception as e:
            log(log_file, f"Error searching buttons: {e}", "WARNING")
        
        # Method 3: Try to find by coordinates (if we know approximate location)
        # This would require manual testing to determine button location
        log(log_file, "Print button not found by text search", "WARNING")
        log(log_file, "You may need to use image recognition or coordinate-based clicking")
        
        return None
        
    except Exception as e:
        log(log_file, f"Error finding print button: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return None


def click_print_button(button, log_file):
    """Click the print button."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 4: Clicking print button")
    log(log_file, "="*60)
    
    if not button:
        log(log_file, "No print button provided", "ERROR")
        return False
    
    try:
        log(log_file, "Checking if button is visible and enabled...")
        if not button.is_visible():
            log(log_file, "Button is not visible", "WARNING")
        if not button.is_enabled():
            log(log_file, "Button is not enabled", "WARNING")
        
        log(log_file, "Clicking print button...")
        button.click_input()
        log(log_file, "✓ Print button clicked")
        time.sleep(2)  # Wait for print dialog to open
        return True
        
    except Exception as e:
        log(log_file, f"Error clicking print button: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def close_preview_window(window, log_file):
    """Close the preview window."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 5: Closing preview window")
    log(log_file, "="*60)
    
    if not window:
        log(log_file, "No window provided", "ERROR")
        return False
    
    try:
        # Method 1: Try to find and click close button
        close_texts = ["Close", "X", "Cancel", "Exit"]
        for text in close_texts:
            try:
                close_button = window.child_window(title=text, control_type="Button")
                if close_button.exists():
                    log(log_file, f"Found close button by text: '{text}'")
                    close_button.click_input()
                    log(log_file, "✓ Close button clicked")
                    time.sleep(1)
                    return True
            except:
                pass
        
        # Method 2: Try to close window directly
        log(log_file, "Close button not found, attempting to close window directly...")
        try:
            window.close()
            log(log_file, "✓ Window closed directly")
            time.sleep(1)
            return True
        except Exception as e:
            log(log_file, f"Could not close window directly: {e}", "WARNING")
        
        # Method 3: Send Alt+F4
        log(log_file, "Attempting Alt+F4...")
        try:
            window.type_keys("%{F4}")  # Alt+F4
            log(log_file, "✓ Sent Alt+F4")
            time.sleep(1)
            return True
        except Exception as e:
            log(log_file, f"Could not send Alt+F4: {e}", "WARNING")
        
        log(log_file, "Could not close window", "ERROR")
        return False
        
    except Exception as e:
        log(log_file, f"Error closing window: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def main():
    """Main test function."""
    log_file = setup_logging()
    
    log(log_file, "="*60)
    log(log_file, "Preview Window Test Script")
    log(log_file, "="*60)
    log(log_file, f"Test started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(log_file, "")
    
    if not PYWINAUTO_AVAILABLE:
        log(log_file, "ERROR: pywinauto is not installed", "ERROR")
        log(log_file, "Install with: pip install pywinauto", "ERROR")
        log_file.close()
        return
    
    log(log_file, "IMPORTANT: This script assumes a preview window is already open.")
    log(log_file, "Please open a label preview window before continuing.")
    log(log_file, "")
    input("Press Enter when preview window is open...")
    log(log_file, "")
    
    try:
        # Step 1: List all windows
        list_all_windows(log_file)
        
        # Step 2: Find preview window
        preview_window = find_preview_window(log_file, timeout=10)
        
        if not preview_window:
            log(log_file, "Could not find preview window", "ERROR")
            log(log_file, "Please ensure a preview window is open and try again", "ERROR")
            return
        
        # Step 3: Extract text
        extracted_text = extract_text_from_preview_window(preview_window, log_file)
        if extracted_text:
            log(log_file, f"✓ Successfully extracted text ({len(extracted_text)} chars)")
        else:
            log(log_file, "Could not extract text from preview window", "WARNING")
        
        # Step 4: Find print button
        print_button = find_print_button(preview_window, log_file)
        if print_button:
            log(log_file, "✓ Print button found")
            
            # Ask user if they want to click it
            log(log_file, "")
            response = input("Do you want to click the print button? (y/n): ").strip().lower()
            if response == 'y':
                click_print_button(print_button, log_file)
                log(log_file, "")
                log(log_file, "Waiting 10 seconds for print dialog...")
                time.sleep(10)
        else:
            log(log_file, "Print button not found", "WARNING")
            log(log_file, "You may need to manually identify the button location")
        
        # Step 5: Close window
        log(log_file, "")
        response = input("Do you want to close the preview window? (y/n): ").strip().lower()
        if response == 'y':
            close_preview_window(preview_window, log_file)
        
        log(log_file, "")
        log(log_file, "="*60)
        log(log_file, "Test completed")
        log(log_file, "="*60)
        
    except Exception as e:
        log(log_file, f"Test failed with error: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
    finally:
        log_file.close()
        print(f"\nTest log saved to: {LOG_FILE}")


if __name__ == "__main__":
    main()
