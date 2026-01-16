"""
Test script for Windows file save dialog automation.
This test locates the file save dialog, sets the save location to testing/,
and sets the filename as file_save_test.

Usage:
    1. Open a file save/download dialog in your browser
    2. Run: python test_file_save_dialog.py
    3. The script will automatically fill in the directory and filename, then save
"""

import time
import sys
from pathlib import Path
from datetime import datetime

try:
    from pywinauto import Desktop, Application
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

# Logging setup
LOG_DIR = Path(__file__).parent
LOG_FILE = LOG_DIR / "test_file_save_dialog_log.txt"

# Target save path
TARGET_DIR = LOG_DIR  # testing/ directory
TARGET_FILENAME = "file_save_test"


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


def find_save_dialog_pywinauto(log_file, timeout=30, debug=False):
    """
    Find the Windows file save dialog using pywinauto.
    
    Args:
        log_file: Log file handle
        timeout: Timeout in seconds
        debug: If True, list all visible windows when dialog not found
    
    Returns:
        Window object if found, None otherwise
    """
    if not PYWINAUTO_AVAILABLE:
        log(log_file, "pywinauto not available", "ERROR")
        return None
    
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 1: Finding file save dialog using pywinauto")
    log(log_file, "="*60)
    
    # Windows that should be excluded (editor windows, browsers, etc.)
    excluded_keywords = [
        "cursor", "vscode", "visual studio", "notepad++", "sublime", 
        "fifra-automation", "python", "pycharm", "intellij",
        "microsoft edge", "chrome", "firefox", "opera", "brave"
    ]
    
    # Common dialog window titles and classes
    dialog_titles = [
        "Save As",
        "Save File",
        "Save Picture As",  # Edge browser specific
        "Save Image As",    # Browser specific
        "File Download",    # Browser download dialog
        "Choose File to Save",
    ]
    
    dialog_classes = [
        "#32770",  # Common dialog class (most important - standard Windows file dialog)
        "SaveDialog",
        "FileSaveDialog",
        "FileDownloadDialog",
    ]
    
    start_time = time.time()
    desktop = Desktop(backend="win32")
    all_windows_listed = False
    
    while time.time() - start_time < timeout:
        try:
            windows = desktop.windows()
            potential_dialogs = []
            
            for window in windows:
                try:
                    if not window.is_visible():
                        continue
                    
                    title = window.window_text()
                    class_name = window.class_name()
                    title_lower = title.lower()
                    
                    # Skip excluded windows (editors, browsers themselves, etc.)
                    if any(keyword in title_lower for keyword in excluded_keywords):
                        continue
                    
                    # Priority 1: Check for standard Windows common dialog class (#32770)
                    # This is the most reliable indicator of a file dialog
                    if "#32770" in class_name or class_name == "#32770":
                        # Verify it's actually a file dialog by checking for file dialog controls
                        try:
                            descendants = window.descendants()
                            has_filename_edit = any("Edit" in str(ctrl.control_type()) for ctrl in descendants)
                            has_save_button = any("Save" in str(ctrl.window_text()) and "Button" in str(ctrl.control_type()) for ctrl in descendants)
                            
                            if has_filename_edit or has_save_button:
                                log(log_file, f"✓ Found dialog by class '#32770' (Windows common dialog):")
                                log(log_file, f"    Title: '{title}'")
                                log(log_file, f"    Class: '{class_name}'")
                                return window
                            else:
                                # It's a #32770 but might not be a file dialog
                                potential_dialogs.append((window, title, class_name, "class #32770 (unverified)"))
                        except Exception as e:
                            # Still might be our dialog, add to potentials
                            potential_dialogs.append((window, title, class_name, "class #32770 (check failed)"))
                    
                    # Priority 2: Check other dialog classes
                    for dialog_class in dialog_classes:
                        if dialog_class != "#32770" and dialog_class in class_name:
                            log(log_file, f"✓ Found dialog by class '{dialog_class}':")
                            log(log_file, f"    Title: '{title}'")
                            log(log_file, f"    Class: '{class_name}'")
                            return window
                    
                    # Priority 3: Check for exact title matches (not partial)
                    for dialog_title in dialog_titles:
                        if title == dialog_title or title.lower() == dialog_title.lower():
                            log(log_file, f"✓ Found dialog by exact title '{dialog_title}':")
                            log(log_file, f"    Title: '{title}'")
                            log(log_file, f"    Class: '{class_name}'")
                            return window
                    
                    # Priority 4: Check for title patterns AND verify it has file dialog controls
                    for dialog_title in dialog_titles:
                        if dialog_title.lower() in title_lower:
                            # Verify it has file dialog-like controls
                            try:
                                descendants = window.descendants()
                                has_combo = any("ComboBox" in str(ctrl.control_type()) for ctrl in descendants)
                                has_edit = any("Edit" in str(ctrl.control_type()) for ctrl in descendants)
                                has_button = any("Button" in str(ctrl.control_type()) for ctrl in descendants)
                                
                                if has_edit and has_button:  # File dialogs always have Edit and Button
                                    log(log_file, f"✓ Found dialog by title pattern '{dialog_title}' with verified controls:")
                                    log(log_file, f"    Title: '{title}'")
                                    log(log_file, f"    Class: '{class_name}'")
                                    return window
                                else:
                                    potential_dialogs.append((window, title, class_name, f"title '{dialog_title}' (no dialog controls)"))
                            except:
                                potential_dialogs.append((window, title, class_name, f"title '{dialog_title}' (check failed)"))
                    
                    # Priority 5: Look for windows with file dialog controls even without matching title
                    try:
                        descendants = window.descendants()
                        edits = [ctrl for ctrl in descendants if "Edit" in str(ctrl.control_type())]
                        buttons = [ctrl for ctrl in descendants if "Button" in str(ctrl.control_type())]
                        combos = [ctrl for ctrl in descendants if "ComboBox" in str(ctrl.control_type())]
                        
                        # File dialogs typically have at least one Edit, one Button (Save/Cancel), and optionally a ComboBox
                        if len(edits) > 0 and len(buttons) >= 2 and len(combos) >= 1:
                            # Check if buttons include Save or similar
                            button_texts = [str(btn.window_text()).lower() for btn in buttons]
                            if any(text in ["save", "ok", "open"] for text in button_texts):
                                log(log_file, f"✓ Found dialog by file dialog controls:")
                                log(log_file, f"    Title: '{title}'")
                                log(log_file, f"    Class: '{class_name}'")
                                log(log_file, f"    Has {len(edits)} Edit(s), {len(buttons)} Button(s), {len(combos)} ComboBox(es)")
                                return window
                    except:
                        pass
                    
                except Exception as e:
                    continue
            
            time.sleep(0.5)
            
        except Exception as e:
            log(log_file, f"Error during search: {e}", "WARNING")
            time.sleep(0.5)
    
    # If we didn't find the dialog and debug is enabled, list all windows
    if debug or not all_windows_listed:
        log(log_file, f"\nFile save dialog not found after {timeout} seconds", "ERROR")
        log(log_file, "\nDEBUG: Listing all visible windows:")
        log(log_file, "="*60)
        try:
            desktop = Desktop(backend="win32")
            windows = desktop.windows()
            for i, window in enumerate(windows, 1):
                try:
                    if window.is_visible():
                        title = window.window_text()
                        class_name = window.class_name()
                        log(log_file, f"{i}. Title: '{title}' | Class: '{class_name}'")
                except:
                    pass
        except Exception as e:
            log(log_file, f"Error listing windows: {e}", "WARNING")
    
    return None


def handle_dialog_pywinauto(dialog, save_path, log_file):
    """
    Handle the file save dialog using pywinauto.
    
    Args:
        dialog: Dialog window object
        save_path: Path object for the target file
        log_file: Log file handle
    
    Returns:
        bool: True if successful
    """
    if not PYWINAUTO_AVAILABLE:
        return False
    
    try:
        log(log_file, "\n" + "="*60)
        log(log_file, "STEP 2: Handling dialog using pywinauto")
        log(log_file, "="*60)
        
        # Focus the dialog
        dialog.set_focus()
        time.sleep(0.5)
        
        # Try to find and set the filename field
        # Common patterns: "Edit" control for filename, "ComboBox" for path
        
        # Method 1: Try to find filename Edit control
        try:
            # Get all child controls
            filename_edit = None
            path_combo = None
            
            # Try different ways to find the controls
            try:
                filename_edit = dialog.child_window(control_type="Edit", found_index=0)
            except:
                pass
            
            try:
                # Sometimes filename is the last Edit control
                edits = [ctrl for ctrl in dialog.descendants(control_type="Edit")]
                if edits:
                    filename_edit = edits[-1]  # Usually filename is last
            except:
                pass
            
            # Try to find path ComboBox
            try:
                path_combo = dialog.child_window(control_type="ComboBox", found_index=0)
            except:
                pass
            
            # Set the path first if ComboBox found
            if path_combo:
                try:
                    directory_path = str(save_path.parent.absolute())
                    log(log_file, f"Setting directory path: {directory_path}")
                    path_combo.set_edit_text(directory_path)
                    time.sleep(0.5)
                    path_combo.type_keys("{ENTER}")
                    time.sleep(1)
                except Exception as e:
                    log(log_file, f"Could not set path via ComboBox: {e}", "WARNING")
            
            # Set the filename
            if filename_edit:
                try:
                    filename = save_path.name
                    log(log_file, f"Setting filename: {filename}")
                    filename_edit.set_edit_text(filename)
                    time.sleep(0.5)
                    log(log_file, "✓ Filename set using Edit control")
                except Exception as e:
                    log(log_file, f"Could not set filename via Edit: {e}", "WARNING")
                    filename_edit = None
            
            # If Edit method didn't work, try typing directly
            if not filename_edit:
                log(log_file, "Trying direct keyboard input...")
                dialog.type_keys(f"{save_path.parent.absolute()}\\{save_path.name}")
                time.sleep(0.5)
            
            # Click Save button or press Enter
            try:
                save_button = dialog.child_window(title="Save", control_type="Button")
                if save_button.exists():
                    save_button.click()
                    log(log_file, "✓ Clicked Save button")
                    return True
            except:
                pass
            
            # Fallback: Press Enter
            dialog.type_keys("{ENTER}")
            log(log_file, "✓ Pressed Enter to save")
            time.sleep(2)
            return True
            
        except Exception as e:
            log(log_file, f"Error in pywinauto method: {e}", "WARNING")
            import traceback
            log(log_file, traceback.format_exc(), "WARNING")
            return False
            
    except Exception as e:
        log(log_file, f"Error handling dialog with pywinauto: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def handle_dialog_pyautogui(save_path, log_file):
    """
    Handle the file save dialog using pyautogui (fallback method).
    
    Args:
        save_path: Path object for the target file
        log_file: Log file handle
    
    Returns:
        bool: True if successful
    """
    if not PYAutoGUI_AVAILABLE:
        log(log_file, "pyautogui not available", "ERROR")
        return False
    
    try:
        log(log_file, "\n" + "="*60)
        log(log_file, "STEP 2 (FALLBACK): Handling dialog using pyautogui")
        log(log_file, "="*60)
        
        # Wait a moment for dialog to be ready
        time.sleep(1)
        
        # Method 1: Use Alt+D to focus address bar, then type path
        log(log_file, "Attempting to set directory path using Alt+D...")
        pyautogui.hotkey('alt', 'd')
        time.sleep(0.5)
        
        # Type the directory path
        directory_path = str(save_path.parent.absolute())
        log(log_file, f"Typing directory path: {directory_path}")
        pyautogui.typewrite(directory_path, interval=0.05)
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(1.5)  # Wait for directory to change
        
        # Navigate to filename field (usually Tab or just click)
        # Filename field is typically already focused, but try Tab just in case
        pyautogui.press('tab')
        time.sleep(0.3)
        
        # Clear existing text (Ctrl+A then type new filename)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.2)
        
        # Type the filename
        filename = save_path.name
        log(log_file, f"Typing filename: {filename}")
        pyautogui.typewrite(filename, interval=0.05)
        time.sleep(0.5)
        
        # Press Enter to save
        pyautogui.press('enter')
        log(log_file, "✓ Pressed Enter to save file")
        
        # Wait for save to complete
        time.sleep(2)
        
        # Verify file was saved
        if save_path.exists():
            log(log_file, f"✓ File saved successfully: {save_path}")
            return True
        else:
            log(log_file, f"WARNING: File not found after save: {save_path}", "WARNING")
            log(log_file, "File may still be saving or path may be different", "WARNING")
            return False
        
    except Exception as e:
        log(log_file, f"Error handling dialog with pyautogui: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def main():
    """Main test function."""
    log_file = setup_logging()
    
    log(log_file, "="*60)
    log(log_file, "File Save Dialog Automation Test")
    log(log_file, "="*60)
    log(log_file, f"Test started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(log_file, "")
    
    # Check dependencies
    if not PYWINAUTO_AVAILABLE and not PYAutoGUI_AVAILABLE:
        log(log_file, "ERROR: Neither pywinauto nor pyautogui is installed", "ERROR")
        log(log_file, "Please install at least one: pip install pywinauto pyautogui", "ERROR")
        log_file.close()
        return
    
    # Set target save path
    target_path = TARGET_DIR / TARGET_FILENAME
    log(log_file, f"Target save directory: {TARGET_DIR}")
    log(log_file, f"Target filename: {TARGET_FILENAME}")
    log(log_file, f"Full target path: {target_path}")
    log(log_file, "")
    log(log_file, "IMPORTANT: Please open a file save/download dialog in your browser")
    log(log_file, "before continuing. The script will wait up to 30 seconds for the dialog.")
    log(log_file, "")
    
    input("Press Enter when the file save dialog is open...")
    log(log_file, "")
    
    try:
        success = False
        
        # Try pywinauto method first (more reliable)
        if PYWINAUTO_AVAILABLE:
            dialog = find_save_dialog_pywinauto(log_file, timeout=30, debug=True)
            
            if dialog:
                success = handle_dialog_pywinauto(dialog, target_path, log_file)
                
                if success:
                    log(log_file, "✓ File dialog handled successfully using pywinauto")
                    time.sleep(1)
                    # Verify file was saved
                    if target_path.exists():
                        log(log_file, f"✓ File confirmed saved: {target_path}")
                    else:
                        log(log_file, f"⚠ File not yet found (may still be saving): {target_path}", "WARNING")
            else:
                log(log_file, "Dialog not found - debug information has been logged above", "WARNING")
        
        # Fallback to pyautogui if pywinauto didn't work
        if not success and PYAutoGUI_AVAILABLE:
            log(log_file, "")
            log(log_file, "pywinauto method didn't work, trying pyautogui fallback...")
            success = handle_dialog_pyautogui(target_path, log_file)
            
            if success:
                log(log_file, "✓ File dialog handled successfully using pyautogui")
            else:
                log(log_file, "⚠ File dialog handling may have failed", "WARNING")
                log(log_file, "Please verify the file was saved manually", "WARNING")
        
        if not success:
            log(log_file, "ERROR: Could not handle file dialog", "ERROR")
            log(log_file, "Troubleshooting:")
            log(log_file, "1. Ensure the file save dialog is open and visible")
            log(log_file, "2. Check if the dialog title matches common patterns")
            log(log_file, "3. Try manually saving once to see the dialog structure")
            log(log_file, "4. Check the log file for more details")
        
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
