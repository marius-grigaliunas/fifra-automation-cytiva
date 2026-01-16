"""
Test script for preview window interaction using screenshots and image matching.
This test locates the preview window, uses image template matching to find icon buttons,
clicks them using coordinates, and handles the Windows file dialog using pyautogui.

Usage:
    python test_preview_window.py
"""

import time
import sys
from pathlib import Path
from datetime import datetime
import ctypes
from ctypes import wintypes

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
    import numpy as np
    PILLOW_AVAILABLE = True
    NUMPY_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False
    NUMPY_AVAILABLE = False
    print("WARNING: Pillow or numpy not installed. Install with: pip install Pillow numpy")

try:
    import cv2
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False
    print("WARNING: opencv-python not installed. Install with: pip install opencv-python")
    print("Will use coordinate-based approach as fallback")

# Logging setup
LOG_DIR = Path(__file__).parent
LOG_FILE = LOG_DIR / "test_preview_window_log.txt"

# Default button positions (relative to window, in pixels from top-left)
# These can be adjusted based on your preview window layout
# Format: (offset_from_left, offset_from_top)
# Note: Coordinates are relative to the WINDOW, not the browser chrome
# The enLabel toolbar is typically below the browser address bar (around 150-200px from top)
DEFAULT_BUTTON_POSITIONS = {
    'print': (50, 180),  # Print button is leftmost in enLabel toolbar (below browser chrome)
}


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


def get_monitor_info():
    """Get information about all monitors using Windows API."""
    monitors = []
    
    def monitor_enum_proc(hmonitor, hdc, lprect, lparam):
        """Callback for EnumDisplayMonitors."""
        rect = lprect.contents
        monitors.append({
            'handle': hmonitor,
            'left': rect.left,
            'top': rect.top,
            'right': rect.right,
            'bottom': rect.bottom,
            'width': rect.right - rect.left,
            'height': rect.bottom - rect.top
        })
        return True
    
    # Define MONITORENUMPROC callback type
    MONITORENUMPROC = ctypes.WINFUNCTYPE(
        ctypes.c_bool,
        ctypes.c_ulong,
        ctypes.c_ulong,
        ctypes.POINTER(wintypes.RECT),
        ctypes.c_ulong
    )
    
    callback = MONITORENUMPROC(monitor_enum_proc)
    
    try:
        ctypes.windll.user32.EnumDisplayMonitors(None, None, callback, 0)
    except Exception as e:
        print(f"Error enumerating monitors: {e}")
    
    return monitors


def detect_window_monitor(window, log_file):
    """
    Detect which monitor the window is on.
    
    Returns:
        dict: Monitor info containing 'left', 'top', 'right', 'bottom', 'width', 'height'
              or None if detection fails
    """
    try:
        monitors = get_monitor_info()
        log(log_file, f"Found {len(monitors)} monitor(s)")
        
        # Get window rectangle (screen coordinates)
        window_rect = window.rectangle()
        window_center_x = window_rect.left + (window_rect.width() // 2)
        window_center_y = window_rect.top + (window_rect.height() // 2)
        
        log(log_file, f"Window center: ({window_center_x}, {window_center_y})")
        
        # Find which monitor contains the window center
        for i, monitor in enumerate(monitors):
            log(log_file, f"Monitor {i+1}: left={monitor['left']}, top={monitor['top']}, "
                f"right={monitor['right']}, bottom={monitor['bottom']}, "
                f"width={monitor['width']}, height={monitor['height']}")
            
            if (monitor['left'] <= window_center_x <= monitor['right'] and
                monitor['top'] <= window_center_y <= monitor['bottom']):
                log(log_file, f"✓ Window is on Monitor {i+1}")
                return monitor
        
        # If no monitor found, return primary monitor (first one)
        log(log_file, "Window center not found in any monitor, using primary monitor", "WARNING")
        return monitors[0] if monitors else None
        
    except Exception as e:
        log(log_file, f"Error detecting monitor: {e}", "ERROR")
        return None


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
            
            time.sleep(0.2)
            
        except Exception as e:
            log(log_file, f"Error during search: {e}", "WARNING")
            time.sleep(0.2)
    
    log(log_file, f"Preview window not found after {timeout} seconds", "ERROR")
    return None


def capture_window_screenshot(window, log_file, monitor=None):
    """
    Capture a screenshot of the window.
    
    Args:
        window: Window object from pywinauto
        log_file: Log file handle
        monitor: Optional monitor info dict for multi-monitor support
    
    Returns:
        PIL Image object or None if capture fails
    """
    if not PILLOW_AVAILABLE:
        log(log_file, "Pillow not available", "ERROR")
        return None
    
    try:
        # Get window rectangle (screen coordinates)
        window_rect = window.rectangle()
        log(log_file, f"Window rectangle: left={window_rect.left}, top={window_rect.top}, "
            f"width={window_rect.width()}, height={window_rect.height()}")
        
        screenshot_left = window_rect.left
        screenshot_top = window_rect.top
        screenshot_width = window_rect.width()
        screenshot_height = window_rect.height()
        
        # Capture screenshot using pyautogui
        screenshot = pyautogui.screenshot(region=(
            screenshot_left,
            screenshot_top,
            screenshot_width,
            screenshot_height
        ))
        
        log(log_file, f"✓ Screenshot captured: {screenshot_width}x{screenshot_height} pixels")
        
        # Save screenshot for debugging
        screenshot_path = LOG_DIR / f"preview_window_screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        screenshot.save(screenshot_path)
        log(log_file, f"Screenshot saved to: {screenshot_path}")
        
        return screenshot
        
    except Exception as e:
        log(log_file, f"Error capturing screenshot: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return None


def detect_button_regions_with_opencv(screenshot, log_file):
    """
    Use OpenCV to detect button-like regions in the enLabel application toolbar.
    This searches BELOW the browser chrome to find buttons in the preview application.
    
    Args:
        screenshot: PIL Image object
        log_file: Log file handle
    
    Returns:
        list of dicts with 'center_x', 'center_y', 'width', 'height' or empty list
    """
    if not CV2_AVAILABLE or not NUMPY_AVAILABLE:
        log(log_file, "OpenCV not available, skipping button detection", "WARNING")
        return []
    
    try:
        log(log_file, "Attempting to detect button regions with OpenCV...")
        log(log_file, "Searching in enLabel application toolbar area (below browser chrome)...")
        
        # Convert PIL image to numpy array for OpenCV
        img_array = np.array(screenshot)
        img_gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
        
        # Skip browser chrome (address bar, IE mode button, etc.)
        # Browser chrome is typically 100-150px tall
        # enLabel header is ~30-50px below that
        # enLabel toolbar with print button is ~30-50px below the header
        # So search area starts around 150-200px from top
        browser_chrome_height = 150  # Skip browser address bar and controls
        toolbar_search_top = browser_chrome_height
        toolbar_search_bottom = browser_chrome_height + 80  # enLabel toolbar is ~80px tall
        
        # Ensure we don't go beyond image bounds
        toolbar_search_top = min(toolbar_search_top, img_gray.shape[0])
        toolbar_search_bottom = min(toolbar_search_bottom, img_gray.shape[0])
        
        if toolbar_search_bottom <= toolbar_search_top:
            log(log_file, f"Screenshot too small for toolbar detection (height: {img_gray.shape[0]}px)", "WARNING")
            return []
        
        log(log_file, f"Searching for buttons in region: y={toolbar_search_top} to {toolbar_search_bottom}")
        toolbar_region = img_gray[toolbar_search_top:toolbar_search_bottom, :]
        
        # Use edge detection to find button boundaries
        edges = cv2.Canny(toolbar_region, 50, 150)
        
        # Find contours
        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        # Filter contours that are roughly button-sized (20-50 pixels wide, 20-40 pixels tall)
        button_regions = []
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            # Adjust y coordinate to be relative to full screenshot (not just toolbar region)
            y_absolute = y + toolbar_search_top
            
            if 15 <= w <= 60 and 15 <= h <= 45:  # Typical button size range
                center_x = x + (w // 2)
                center_y = y_absolute + (h // 2)
                button_regions.append({
                    'center_x': center_x,
                    'center_y': center_y,
                    'width': w,
                    'height': h,
                    'left': x,
                    'top': y_absolute
                })
                log(log_file, f"  Found potential button region at ({center_x}, {center_y}), size: {w}x{h}")
        
        # Sort by x position (left to right) - print button should be leftmost
        button_regions.sort(key=lambda b: b['center_x'])
        
        log(log_file, f"Detected {len(button_regions)} potential button regions in enLabel toolbar")
        
        return button_regions
        
    except Exception as e:
        log(log_file, f"Error detecting button regions: {type(e).__name__}: {str(e)}", "WARNING")
        return []


def get_button_coordinates(window, button_name, log_file, detected_regions=None):
    """
    Get coordinates for the print button, either from detected regions, absolute coordinates,
    or default positions.
    
    Args:
        window: Window object from pywinauto
        button_name: Name of button ('print' only - save button does not exist)
        log_file: Log file handle
        detected_regions: Optional list of detected button regions from OpenCV
    
    Returns:
        dict with 'center_x', 'center_y', 'use_absolute' flag, or None
    """
    try:
        if button_name != 'print':
            log(log_file, f"Only 'print' button is supported (button '{button_name}' requested)", "WARNING")
            return None
        
        window_rect = window.rectangle()
        
        # If we have detected regions, use the leftmost one (print button is first)
        if detected_regions:
            # Print button should be the leftmost button in the enLabel toolbar
            region = detected_regions[0]
            log(log_file, f"Using detected leftmost region for 'print' button: "
                f"({region['center_x']}, {region['center_y']})")
            # Detected regions are relative to window
            region['use_absolute'] = False
            return region
        
        # Known working absolute coordinates for print button
        # These are screen coordinates that work reliably
        ABSOLUTE_PRINT_COORDS = (493, 421)
        
        # Calculate relative coordinates from absolute using current window position
        abs_x, abs_y = ABSOLUTE_PRINT_COORDS
        rel_x = abs_x - window_rect.left
        rel_y = abs_y - window_rect.top
        
        log(log_file, f"Using known print button coordinates (absolute: {abs_x}, {abs_y})")
        log(log_file, f"Window position: ({window_rect.left}, {window_rect.top})")
        log(log_file, f"Calculated relative coordinates: ({rel_x}, {rel_y})")
        
        return {
            'center_x': abs_x,
            'center_y': abs_y,
            'left': abs_x - 10,
            'top': abs_y - 10,
            'width': 20,
            'height': 20,
            'use_absolute': True  # Use absolute coordinates directly
        }
        
    except Exception as e:
        log(log_file, f"Error getting button coordinates: {type(e).__name__}: {str(e)}", "ERROR")
        return None


def click_at_coordinates(window, button_info, log_file, monitor=None):
    """
    Click at the specified coordinates (absolute or relative to window).
    
    Args:
        window: Window object from pywinauto
        button_info: Dict with 'center_x', 'center_y' and optional 'use_absolute' flag
        log_file: Log file handle
        monitor: Optional monitor info for coordinate adjustment
    
    Returns:
        bool: True if click was successful
    """
    if not PYAutoGUI_AVAILABLE:
        log(log_file, "pyautogui not available", "ERROR")
        return False
    
    try:
        log(log_file, "\n" + "="*60)
        log(log_file, "STEP 5: Clicking print button at coordinates")
        log(log_file, "="*60)
        
        # Get window rectangle (screen coordinates)
        window_rect = window.rectangle()
        
        # Determine if coordinates are absolute or relative
        use_absolute = button_info.get('use_absolute', False)
        
        if use_absolute:
            # Use absolute coordinates directly
            absolute_x = button_info['center_x']
            absolute_y = button_info['center_y']
            log(log_file, f"Using absolute screen coordinates: ({absolute_x}, {absolute_y})")
        else:
            # Calculate absolute screen coordinates from relative
            absolute_x = window_rect.left + button_info['center_x']
            absolute_y = window_rect.top + button_info['center_y']
            log(log_file, f"Window position: ({window_rect.left}, {window_rect.top})")
            log(log_file, f"Button relative position: ({button_info['center_x']}, {button_info['center_y']})")
            log(log_file, f"Absolute click coordinates: ({absolute_x}, {absolute_y})")
        
        # Focus the window first
        try:
            window.set_focus()
            time.sleep(0.1)
        except Exception as e:
            log(log_file, f"Could not focus window: {e}", "WARNING")
        
        # Click at the coordinates using pyautogui
        pyautogui.click(absolute_x, absolute_y)
        log(log_file, f"✓ Clicked print button at ({absolute_x}, {absolute_y})")
        
        time.sleep(1)  # Wait for printer dialog to open
        
        return True
        
    except Exception as e:
        log(log_file, f"Error clicking coordinates: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def handle_printer_dialog(log_file):
    """
    Handle the Windows printer selection dialog.
    Selects "Microsoft Print to PDF" and presses OK.
    
    Args:
        log_file: Log file handle
    
    Returns:
        bool: True if dialog was handled successfully
    """
    if not PYAutoGUI_AVAILABLE:
        log(log_file, "pyautogui not available", "ERROR")
        return False
    
    try:
        log(log_file, "\n" + "="*60)
        log(log_file, "STEP 6: Handling printer selection dialog")
        log(log_file, "="*60)
        
        # Wait for printer dialog to appear
        time.sleep(1)
        
        # Try to find the printer dialog window using pywinauto
        if PYWINAUTO_AVAILABLE:
            try:
                desktop = Desktop(backend="win32")
                printer_dialogs = [
                    w for w in desktop.windows()
                    if w.is_visible() and (
                        "print" in w.window_text().lower() or
                        "printer" in w.window_text().lower() or
                        w.class_name() == "#32770"  # Common dialog class
                    )
                ]
                
                if printer_dialogs:
                    printer_dialog = printer_dialogs[0]
                    log(log_file, f"Found printer dialog: '{printer_dialog.window_text()}'")
                    printer_dialog.set_focus()
                    time.sleep(0.2)
            except Exception as e:
                log(log_file, f"Could not find printer dialog window: {e}", "WARNING")
                log(log_file, "Continuing with keyboard navigation...")
        
        # Method 1: Type the printer name directly (often works if combo box is focused)
        log(log_file, "Typing 'Microsoft Print to PDF' to select printer...")
        pyautogui.typewrite("Microsoft Print to PDF", interval=0.05)
        time.sleep(0.5)  # Wait for autocomplete/selection
        
        # Press Enter to confirm the printer selection from dropdown
        pyautogui.press('enter')
        time.sleep(0.3)
        
        # Navigate to OK button and press it
        # Windows printer dialogs typically use Alt+O for OK button
        log(log_file, "Pressing Alt+O to click OK button...")
        pyautogui.hotkey('alt', 'o')
        time.sleep(0.5)
        
        # Alternative: If Alt+O didn't work, try Tab to navigate to OK and press Enter
        # But first, try Enter one more time in case OK is already focused
        log(log_file, "Pressing Enter to confirm (OK button may be focused)...")
        pyautogui.press('enter')
        
        # Wait for dialog to close
        time.sleep(1)
        
        log(log_file, "✓ Printer selection confirmed (dialog should be closed)")
        return True
        
    except Exception as e:
        log(log_file, f"Error handling printer dialog: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def resize_preview_window(log_file, start_x=850, start_y=827, end_x=1450, end_y=868):
    """
    Resize the preview window by dragging the bottom-right corner.
    
    Args:
        log_file: Log file handle
        start_x: Starting X coordinate (bottom-right corner)
        start_y: Starting Y coordinate (bottom-right corner)
        end_x: Ending X coordinate (drag to position)
        end_y: Ending Y coordinate (drag to position)
    
    Returns:
        bool: True if resize was successful
    """
    if not PYAutoGUI_AVAILABLE:
        log(log_file, "pyautogui not available", "ERROR")
        return False
    
    try:
        log(log_file, "\n" + "="*60)
        log(log_file, "STEP 2: Resizing preview window")
        log(log_file, "="*60)
        
        log(log_file, f"Dragging window from bottom-right corner ({start_x}, {start_y}) to ({end_x}, {end_y})")
        
        # Drag from starting position to end position
        # pyautogui.drag() moves relative to current position and holds mouse button
        # Move to starting position first
        pyautogui.moveTo(start_x, start_y)
        time.sleep(0.1)
        
        # Drag to end position - drag() handles mouse down/up automatically
        drag_x = end_x - start_x
        drag_y = end_y - start_y
        pyautogui.drag(drag_x, drag_y, duration=0.5, button='left')
        
        log(log_file, f"✓ Window resize completed (dragged from ({start_x}, {start_y}) to ({end_x}, {end_y}))")
        
        # Wait for window to finish resizing
        time.sleep(0.5)
        
        return True
        
    except Exception as e:
        log(log_file, f"Error resizing window: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def click_zoom_in_button(log_file, x=540, y=420, times=3):
    """
    Click the zoom in button multiple times.
    
    Args:
        log_file: Log file handle
        x: X coordinate of zoom in button
        y: Y coordinate of zoom in button
        times: Number of times to click the button
    
    Returns:
        bool: True if clicks were successful
    """
    if not PYAutoGUI_AVAILABLE:
        log(log_file, "pyautogui not available", "ERROR")
        return False
    
    try:
        log(log_file, "\n" + "="*60)
        log(log_file, f"STEP 3: Clicking zoom in button {times} times")
        log(log_file, "="*60)
        
        log(log_file, f"Zoom in button coordinates: ({x}, {y})")
        
        # Click the zoom in button the specified number of times
        for i in range(times):
            pyautogui.click(x, y)
            log(log_file, f"  Click {i+1}/{times} at ({x}, {y})")
            time.sleep(0.2)  # Wait between clicks (reduced from 0.5)
        
        log(log_file, f"✓ Clicked zoom in button {times} times")
        
        # Wait for zoom to complete
        time.sleep(0.3)
        
        return True
        
    except Exception as e:
        log(log_file, f"Error clicking zoom in button: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return False


def close_preview_window(window, log_file):
    """Close the preview window."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 7: Closing preview window")
    log(log_file, "="*60)
    
    if not window:
        log(log_file, "No window provided", "ERROR")
        return False
    
    try:
        window_title = window.window_text()
        log(log_file, f"Closing window: '{window_title}'")
        
        # Try to close using window.close()
        window.close()
        log(log_file, "✓ Window.close() called")
        time.sleep(1)
        
        # Verify window is closed
        try:
            if not window.exists():
                log(log_file, "✓ Window successfully closed")
                return True
        except:
            # Window object may be invalid after close - this is success
            log(log_file, "✓ Window closed (object invalidated)")
            return True
        
        # If close() didn't work, try Alt+F4
        window.set_focus()
        time.sleep(0.2)
        window.type_keys("%{F4}")  # Alt+F4
        log(log_file, "✓ Sent Alt+F4")
        time.sleep(1)
        return True
        
    except Exception as e:
        log(log_file, f"Error closing window: {type(e).__name__}: {str(e)}", "ERROR")
        return False


def main():
    """Main test function."""
    log_file = setup_logging()
    
    log(log_file, "="*60)
    log(log_file, "Preview Window Test Script (Image Matching + Coordinate Approach)")
    log(log_file, "="*60)
    log(log_file, f"Test started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(log_file, "")
    
    # Check dependencies
    if not PYWINAUTO_AVAILABLE:
        log(log_file, "ERROR: pywinauto is not installed", "ERROR")
        log_file.close()
        return
    
    if not PYAutoGUI_AVAILABLE:
        log(log_file, "ERROR: pyautogui is not installed", "ERROR")
        log_file.close()
        return
    
    if not PILLOW_AVAILABLE:
        log(log_file, "ERROR: Pillow is not installed", "ERROR")
        log_file.close()
        return
    
    log(log_file, "IMPORTANT: This script assumes a preview window is already open.")
    log(log_file, "Please open a label preview window before continuing.")
    log(log_file, "")
    log(log_file, "NOTE: This script will resize the window, zoom in, then click the PRINT button.")
    log(log_file, "The button is in the enLabel application toolbar (below the browser chrome).")
    log(log_file, "")
    input("Press Enter when preview window is open...")
    log(log_file, "")
    
    try:
        # Step 1: Find preview window
        preview_window = find_preview_window(log_file, timeout=10)
        
        if not preview_window:
            log(log_file, "Could not find preview window", "ERROR")
            log(log_file, "Please ensure a preview window is open and try again", "ERROR")
            return
        
        # Step 2: Resize the preview window (skip screenshot/detection - using absolute coordinates)
        resize_success = resize_preview_window(log_file, start_x=850, start_y=827, end_x=1450, end_y=868)
        
        if not resize_success:
            log(log_file, "Failed to resize window", "WARNING")
            log(log_file, "Continuing anyway...", "WARNING")
        
        # Step 3: Click zoom in button 3 times
        zoom_success = click_zoom_in_button(log_file, x=540, y=420, times=3)
        
        if not zoom_success:
            log(log_file, "Failed to click zoom in button", "WARNING")
            log(log_file, "Continuing anyway...", "WARNING")
        
        # Step 4: Get button coordinates for Print button (using absolute coordinates)
        # Note: Only print button exists - no save button in the preview
        button_name = 'print'
        button_info = get_button_coordinates(preview_window, button_name, log_file, detected_regions=None)
        
        if not button_info:
            log(log_file, "Could not determine button coordinates", "ERROR")
            log(log_file, "Please adjust DEFAULT_BUTTON_POSITIONS in the script", "ERROR")
            log(log_file, "You can find button positions by:")
            log(log_file, "1. Taking a screenshot (saved in testing folder)")
            log(log_file, "2. Opening it in an image editor to find pixel coordinates")
            log(log_file, "3. Updating DEFAULT_BUTTON_POSITIONS with (x, y) offsets from window top-left")
            return
        
        # Step 5: Click the print button
        click_success = click_at_coordinates(preview_window, button_info, log_file, monitor=None)
        
        if not click_success:
            log(log_file, "Failed to click print button", "ERROR")
            return
        
        # Step 6: Handle printer selection dialog
        dialog_success = handle_printer_dialog(log_file)
        
        if not dialog_success:
            log(log_file, "Printer dialog handling may have failed", "WARNING")
            log(log_file, "Please verify the printer was selected manually", "WARNING")
        log(log_file, "")
        response = input("Do you want to close the preview window? (y/n): ").strip().lower()
        if response == 'y':
            close_preview_window(preview_window, log_file)
        
        log(log_file, "")
        log(log_file, "="*60)
        log(log_file, "Test completed")
        log(log_file, "="*60)
        log(log_file, "")
        log(log_file, "TROUBLESHOOTING:")
        log(log_file, "If the print button was not clicked correctly:")
        log(log_file, "1. The script uses absolute coordinates (493, 421) for the print button")
        log(log_file, "2. If your window is in a different position, update ABSOLUTE_PRINT_COORDS")
        log(log_file, "   in the get_button_coordinates() function")
        log(log_file, "")
        log(log_file, "If the printer dialog did not work:")
        log(log_file, "1. The script types 'Microsoft Print to PDF' and presses Enter")
        log(log_file, "2. Make sure 'Microsoft Print to PDF' is installed on your system")
        log(log_file, "3. The script will try Alt+O and Enter to confirm the selection")
        
    except Exception as e:
        log(log_file, f"Test failed with error: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
    finally:
        log_file.close()
        print(f"\nTest log saved to: {LOG_FILE}")


if __name__ == "__main__":
    main()
