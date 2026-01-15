"""
Test script for complete print and save flow.
Tests full flow: preview → print → printer selection → save dialog.

This script assumes a preview window is already open. You should run test_label_search.py
first to open a preview window, then run this script.

Alternatively, you can manually open a preview window before running this script.

Usage:
    python test_print_save_flow.py [--output-dir path/to/output]
    
Example:
    python test_print_save_flow.py --output-dir C:/temp/labels
"""

import time
import sys
from pathlib import Path
from datetime import datetime
import argparse

try:
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import ElementNotFoundError
    import pyautogui
    AUTOMATION_AVAILABLE = True
except ImportError:
    AUTOMATION_AVAILABLE = False
    print("WARNING: pywinauto or pyautogui not installed.")
    print("Install with: pip install pywinauto pyautogui")

# Logging setup
LOG_DIR = Path(__file__).parent
LOG_FILE = LOG_DIR / "test_print_save_flow_log.txt"


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


def find_preview_window(log_file, timeout=10):
    """Find the preview window."""
    if not AUTOMATION_AVAILABLE:
        log(log_file, "Automation libraries not available", "ERROR")
        return None
    
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 1: Finding preview window")
    log(log_file, "="*60)
    
    title_patterns = ["Label Preview", "Preview", "Label", "Print Preview", "enLabel"]
    
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
                    for pattern in title_patterns:
                        if pattern.lower() in title.lower():
                            log(log_file, f"✓ Found preview window: '{title}'")
                            return window
                except:
                    continue
            time.sleep(1)
        except Exception as e:
            log(log_file, f"Error during search: {e}", "WARNING")
            time.sleep(1)
    
    log(log_file, f"Preview window not found after {timeout} seconds", "ERROR")
    return None


def find_print_button(window, log_file):
    """Find print button in preview window."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 2: Finding print button")
    log(log_file, "="*60)
    
    if not window:
        log(log_file, "No window provided", "ERROR")
        return None
    
    try:
        # Search by button text
        button_texts = ["Print", "Print...", "Print Label", "Print Preview"]
        for text in button_texts:
            try:
                button = window.child_window(title=text, control_type="Button")
                if button.exists():
                    log(log_file, f"✓ Found print button: '{text}'")
                    return button
            except:
                pass
        
        # Search all buttons
        try:
            buttons = window.descendants(control_type="Button")
            log(log_file, f"Found {len(buttons)} button(s) in window")
            for button in buttons:
                try:
                    button_text = button.window_text()
                    if "print" in button_text.lower():
                        log(log_file, f"  → Found print button: '{button_text}'")
                        return button
                except:
                    pass
        except Exception as e:
            log(log_file, f"Error searching buttons: {e}", "WARNING")
        
        log(log_file, "Print button not found", "ERROR")
        return None
        
    except Exception as e:
        log(log_file, f"Error finding print button: {type(e).__name__}: {str(e)}", "ERROR")
        return None


def click_print_button(button, log_file):
    """Click the print button."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 3: Clicking print button")
    log(log_file, "="*60)
    
    if not button:
        log(log_file, "No print button provided", "ERROR")
        return False
    
    try:
        log(log_file, "Clicking print button...")
        button.click_input()
        log(log_file, "✓ Print button clicked")
        log(log_file, "Waiting 3 seconds for print dialog to open...")
        time.sleep(3)
        return True
    except Exception as e:
        log(log_file, f"Error clicking print button: {type(e).__name__}: {str(e)}", "ERROR")
        return False


def select_printer(log_file, printer_name="Microsoft Print to PDF", timeout=10):
    """Select printer in print dialog."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 4: Selecting printer")
    log(log_file, "="*60)
    log(log_file, f"Target printer: {printer_name}")
    
    if not AUTOMATION_AVAILABLE:
        log(log_file, "Automation libraries not available", "ERROR")
        return False
    
    try:
        # Find print dialog
        log(log_file, "Looking for print dialog...")
        desktop = Desktop(backend="win32")
        
        start_time = time.time()
        print_dialog = None
        
        while time.time() - start_time < timeout:
            windows = desktop.windows()
            for window in windows:
                try:
                    if not window.is_visible():
                        continue
                    title = window.window_text().lower()
                    if "print" in title and ("dialog" in title or "printer" in title or "select" in title):
                        log(log_file, f"✓ Found print dialog: '{window.window_text()}'")
                        print_dialog = window
                        break
                except:
                    continue
            
            if print_dialog:
                break
            time.sleep(0.5)
        
        if not print_dialog:
            log(log_file, "Print dialog not found", "ERROR")
            log(log_file, "Available windows:")
            windows = desktop.windows()
            for window in windows:
                try:
                    if window.is_visible():
                        log(log_file, f"  - '{window.window_text()}'")
                except:
                    pass
            return False
        
        # Try to find printer selection dropdown or list
        log(log_file, "Looking for printer selection...")
        
        # Method 1: Try to find combo box or list box
        try:
            combo_boxes = print_dialog.descendants(control_type="ComboBox")
            log(log_file, f"Found {len(combo_boxes)} combo box(es)")
            for combo in combo_boxes:
                try:
                    combo_text = combo.window_text()
                    log(log_file, f"  Combo box: '{combo_text}'")
                    # Try to select the printer
                    combo.select(printer_name)
                    log(log_file, f"✓ Selected printer: {printer_name}")
                    time.sleep(1)
                    return True
                except Exception as e:
                    log(log_file, f"  Error selecting from combo: {e}", "WARNING")
        except Exception as e:
            log(log_file, f"Error finding combo boxes: {e}", "WARNING")
        
        # Method 2: Try to find list box
        try:
            list_boxes = print_dialog.descendants(control_type="List")
            log(log_file, f"Found {len(list_boxes)} list box(es)")
            for list_box in list_boxes:
                try:
                    items = list_box.item_texts()
                    log(log_file, f"  List items: {items}")
                    if printer_name in items:
                        list_box.select(printer_name)
                        log(log_file, f"✓ Selected printer from list: {printer_name}")
                        time.sleep(1)
                        return True
                except Exception as e:
                    log(log_file, f"  Error selecting from list: {e}", "WARNING")
        except Exception as e:
            log(log_file, f"Error finding list boxes: {e}", "WARNING")
        
        # Method 3: Use pyautogui to click on printer name (if visible)
        log(log_file, "Attempting to use image recognition to find printer...")
        try:
            # This would require a screenshot of the printer name
            # For now, just log that manual intervention may be needed
            log(log_file, "Could not automatically select printer", "WARNING")
            log(log_file, "You may need to manually select the printer", "WARNING")
            return False
        except Exception as e:
            log(log_file, f"Error with image recognition: {e}", "WARNING")
        
        return False
        
    except Exception as e:
        log(log_file, f"Error selecting printer: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def handle_save_dialog(log_file, output_path, timeout=30):
    """Handle Windows file save dialog."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 5: Handling save dialog")
    log(log_file, "="*60)
    log(log_file, f"Output path: {output_path}")
    
    if not AUTOMATION_AVAILABLE:
        log(log_file, "Automation libraries not available", "ERROR")
        return False
    
    try:
        # Find save dialog
        log(log_file, "Looking for save dialog...")
        desktop = Desktop(backend="win32")
        
        start_time = time.time()
        save_dialog = None
        
        while time.time() - start_time < timeout:
            windows = desktop.windows()
            for window in windows:
                try:
                    if not window.is_visible():
                        continue
                    title = window.window_text().lower()
                    if "save" in title or "save as" in title or "print to pdf" in title:
                        log(log_file, f"✓ Found save dialog: '{window.window_text()}'")
                        save_dialog = window
                        break
                except:
                    continue
            
            if save_dialog:
                break
            time.sleep(0.5)
        
        if not save_dialog:
            log(log_file, "Save dialog not found", "ERROR")
            return False
        
        # Find filename input field
        log(log_file, "Looking for filename input field...")
        try:
            # Common patterns for filename input
            filename_inputs = save_dialog.descendants(control_type="Edit")
            log(log_file, f"Found {len(filename_inputs)} edit field(s)")
            
            # Usually the first edit field is the filename
            if filename_inputs:
                filename_input = filename_inputs[0]
                log(log_file, "Entering filename...")
                filename_input.set_text(str(output_path))
                log(log_file, f"✓ Entered filename: {output_path}")
                time.sleep(1)
            else:
                log(log_file, "No edit fields found, trying alternative method...")
                # Try typing the filename directly
                save_dialog.type_keys(str(output_path))
                log(log_file, f"✓ Typed filename: {output_path}")
                time.sleep(1)
        except Exception as e:
            log(log_file, f"Error entering filename: {e}", "WARNING")
            # Fallback: use pyautogui to type
            try:
                pyautogui.write(str(output_path), interval=0.1)
                log(log_file, f"✓ Typed filename using pyautogui: {output_path}")
                time.sleep(1)
            except Exception as e2:
                log(log_file, f"Error with pyautogui: {e2}", "ERROR")
                return False
        
        # Find and click Save button
        log(log_file, "Looking for Save button...")
        try:
            save_buttons = save_dialog.descendants(control_type="Button")
            for button in save_buttons:
                try:
                    button_text = button.window_text().lower()
                    if "save" in button_text and "cancel" not in button_text:
                        log(log_file, f"✓ Found Save button: '{button.window_text()}'")
                        button.click_input()
                        log(log_file, "✓ Clicked Save button")
                        time.sleep(2)
                        return True
                except:
                    pass
            
            # Fallback: try Enter key or Alt+S
            log(log_file, "Save button not found, trying Enter key...")
            save_dialog.type_keys("{ENTER}")
            log(log_file, "✓ Pressed Enter")
            time.sleep(2)
            return True
        except Exception as e:
            log(log_file, f"Error clicking Save button: {type(e).__name__}: {str(e)}", "ERROR")
            return False
        
    except Exception as e:
        log(log_file, f"Error handling save dialog: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def main():
    """Main test function."""
    parser = argparse.ArgumentParser(description="Test print and save flow")
    parser.add_argument('--output-dir', type=str, default=None, help='Output directory for saved file')
    parser.add_argument('--filename', type=str, default="test_label.pdf", help='Filename for saved file')
    
    args = parser.parse_args()
    
    log_file = setup_logging()
    
    log(log_file, "="*60)
    log(log_file, "Print and Save Flow Test Script")
    log(log_file, "="*60)
    log(log_file, f"Test started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(log_file, "")
    
    if not AUTOMATION_AVAILABLE:
        log(log_file, "ERROR: pywinauto or pyautogui is not installed", "ERROR")
        log(log_file, "Install with: pip install pywinauto pyautogui", "ERROR")
        log_file.close()
        return
    
    log(log_file, "IMPORTANT: This script assumes a preview window is already open.")
    log(log_file, "Please open a label preview window before continuing.")
    log(log_file, "")
    input("Press Enter when preview window is open...")
    log(log_file, "")
    
    # Determine output path
    if args.output_dir:
        output_dir = Path(args.output_dir)
    else:
        output_dir = LOG_DIR / "test_output"
    
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / args.filename
    
    log(log_file, f"Output directory: {output_dir}")
    log(log_file, f"Output filename: {args.filename}")
    log(log_file, f"Full output path: {output_path}")
    log(log_file, "")
    
    try:
        # Step 1: Find preview window
        preview_window = find_preview_window(log_file, timeout=10)
        if not preview_window:
            log(log_file, "Could not find preview window", "ERROR")
            log(log_file, "Please ensure a preview window is open and try again", "ERROR")
            return
        
        # Step 2: Find print button
        print_button = find_print_button(preview_window, log_file)
        if not print_button:
            log(log_file, "Could not find print button", "ERROR")
            return
        
        # Step 3: Click print button
        if not click_print_button(print_button, log_file):
            log(log_file, "Failed to click print button", "ERROR")
            return
        
        # Step 4: Select printer
        log(log_file, "")
        response = input("Do you want to automatically select printer? (y/n): ").strip().lower()
        if response == 'y':
            if not select_printer(log_file):
                log(log_file, "Failed to select printer automatically", "WARNING")
                log(log_file, "Please manually select 'Microsoft Print to PDF' and press Enter...")
                input()
        else:
            log(log_file, "Please manually select 'Microsoft Print to PDF' and press Enter...")
            input()
        
        # Step 5: Handle save dialog
        log(log_file, "")
        response = input("Do you want to automatically handle save dialog? (y/n): ").strip().lower()
        if response == 'y':
            if handle_save_dialog(log_file, output_path):
                log(log_file, "✓ File save completed")
                # Verify file was created
                time.sleep(2)
                if output_path.exists():
                    log(log_file, f"✓ File verified: {output_path}")
                    log(log_file, f"  File size: {output_path.stat().st_size} bytes")
                else:
                    log(log_file, f"⚠ File not found at expected location: {output_path}", "WARNING")
            else:
                log(log_file, "Failed to handle save dialog automatically", "WARNING")
                log(log_file, "Please manually save the file and press Enter...")
                input()
        else:
            log(log_file, "Please manually save the file and press Enter...")
            input()
        
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
