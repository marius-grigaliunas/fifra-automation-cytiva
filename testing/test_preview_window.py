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
import ctypes

try:
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import ElementNotFoundError
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False
    print("WARNING: pywinauto not installed. Install with: pip install pywinauto")

try:
    import win32gui
    import win32con
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

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
    
    # Exclude common editor windows that might match "Preview" in title
    # Note: Chrome_WidgetWin_1 can be Edge browser (which we want) OR Cursor/VS Code (which we don't)
    # So we'll check the title to distinguish them
    excluded_title_keywords = ["cursor", "vscode", "visual studio", "notepad++", "sublime", "fifra-automation"]
    
    if title_patterns is None:
        title_patterns = [
            "enLabel Global Services - Work - Microsoft Edge",  # Most specific - Edge preview window
            "enLabel Global Services - Internet Explorer",  # IE preview window
            "enLabel Global Services",  # Generic enLabel window
            "Label Preview",
            "Print Preview",
            "enLabel",
        ]
    
    if class_patterns is None:
        class_patterns = [
            "IEFrame",  # Internet Explorer frame - most reliable
            "Internet Explorer_Server",  # ActiveX control within IE
            "Chrome_WidgetWin_1",  # Microsoft Edge browser window (can be preview)
            "Shell Embedding",
        ]
    
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 1: Finding preview window")
    log(log_file, "="*60)
    log(log_file, f"Searching with timeout: {timeout} seconds")
    log(log_file, f"Title patterns: {title_patterns}")
    log(log_file, f"Class patterns: {class_patterns}")
    log(log_file, "Note: Prioritizing class-based matching to avoid false positives")
    
    start_time = time.time()
    desktop = Desktop(backend="win32")
    
    # First pass: Check class patterns (more reliable)
    log(log_file, "Pass 1: Checking class patterns (most reliable)...")
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
                    
                    # Skip excluded windows by title keywords (but allow Chrome_WidgetWin_1 if it's enLabel)
                    title_lower = title.lower()
                    if any(keyword in title_lower for keyword in excluded_title_keywords):
                        continue
                    
                    # PRIORITY 1: Check class patterns first (most reliable)
                    for pattern in class_patterns:
                        if pattern.lower() in class_name.lower():
                            # Additional verification based on class type
                            if "chrome_widgetwin" in class_name.lower():
                                # For Edge/Chrome windows, must have "enlabel" in title
                                # Check for "microsoft edge" or just "edge" (case insensitive)
                                has_edge = "microsoft edge" in title_lower or ("edge" in title_lower and "microsoft" in title_lower)
                                if "enlabel" in title_lower and has_edge:
                                    log(log_file, f"✓ Found window by class pattern '{pattern}':")
                                    log(log_file, f"    Title: '{title}'")
                                    log(log_file, f"    Class: '{class_name}'")
                                    return window
                            else:
                                # For IEFrame and other classes, check for enLabel/IE/preview
                                if "enlabel" in title_lower or "internet explorer" in title_lower or "preview" in title_lower:
                                    log(log_file, f"✓ Found window by class pattern '{pattern}':")
                                    log(log_file, f"    Title: '{title}'")
                                    log(log_file, f"    Class: '{class_name}'")
                                    return window
                    
                except Exception as e:
                    continue
            
            time.sleep(0.5)
            
        except Exception as e:
            log(log_file, f"Error during search: {e}", "WARNING")
            time.sleep(0.5)
    
    # Second pass: Check title patterns (if class-based search failed)
    log(log_file, "Pass 2: Checking title patterns...")
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            windows = desktop.windows()
            
            for window in windows:
                try:
                    if not window.is_visible():
                        continue
                    
                    title = window.window_text()
                    class_name = window.class_name()
                    
                    # Skip excluded windows by title keywords
                    title_lower = title.lower()
                    if any(keyword in title_lower for keyword in excluded_title_keywords):
                        continue
                    
                    # PRIORITY 2: Check specific title patterns
                    for pattern in title_patterns:
                        if pattern.lower() in title_lower:
                            # Additional verification: if it's Chrome_WidgetWin_1, make sure it's not Cursor/VS Code
                            if "chrome_widgetwin" in class_name.lower():
                                # Must be Edge browser with enLabel, not editor
                                has_edge = "microsoft edge" in title_lower or ("edge" in title_lower and "microsoft" in title_lower)
                                if "enlabel" in title_lower and has_edge:
                                    log(log_file, f"✓ Found window by title pattern '{pattern}':")
                                    log(log_file, f"    Title: '{title}'")
                                    log(log_file, f"    Class: '{class_name}'")
                                    return window
                            else:
                                # For other classes, accept if title matches
                                log(log_file, f"✓ Found window by title pattern '{pattern}':")
                                log(log_file, f"    Title: '{title}'")
                                log(log_file, f"    Class: '{class_name}'")
                                return window
                    
                except Exception as e:
                    continue
            
            time.sleep(0.5)
            
        except Exception as e:
            log(log_file, f"Error during search: {e}", "WARNING")
            time.sleep(0.5)
    
    log(log_file, f"Preview window not found after {timeout} seconds", "ERROR")
    log(log_file, "Available windows:")
    list_all_windows(log_file)
    return None


def get_activex_control(window, log_file):
    """Get the ActiveX control (Internet Explorer_Server) from an IEFrame window."""
    try:
        window_class = window.class_name()
        if "IEFrame" in window_class:
            log(log_file, "IEFrame detected - searching for Internet Explorer_Server...")
            ie_server = window.child_window(class_name="Internet Explorer_Server")
            if ie_server.exists():
                log(log_file, "✓ Found Internet Explorer_Server (ActiveX control)")
                return ie_server
        return None
    except Exception as e:
        log(log_file, f"Could not get ActiveX control: {e}", "WARNING")
        return None


def extract_text_from_preview_window(window, log_file):
    """Extract text from preview window.
    For ActiveX controls, we need to access the Internet Explorer_Server child.
    """
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 2: Extracting text from preview window")
    log(log_file, "="*60)
    
    if not window:
        log(log_file, "No window provided", "ERROR")
        return None
    
    try:
        window_title = window.window_text()
        window_class = window.class_name()
        log(log_file, f"Window: '{window_title}' (Class: {window_class})")
        log(log_file, "Attempting to extract window text...")
        
        # Method 1: Get window text directly
        window_text = ""
        try:
            window_text = window.window_text()
            if window_text and len(window_text) > 0:
                log(log_file, f"Window text (first 500 chars): {window_text[:500]}")
                if len(window_text) > 500:
                    log(log_file, f"... (truncated, total length: {len(window_text)} chars)")
        except Exception as e:
            log(log_file, f"Could not get window text directly: {e}", "WARNING")
        
        # Method 2: For IEFrame, try to get text from ActiveX control
        if "IEFrame" in window_class:
            log(log_file, "IEFrame detected - accessing ActiveX control...")
            activex = get_activex_control(window, log_file)
            if activex:
                try:
                    activex_text = activex.window_text()
                    if activex_text and len(activex_text) > len(window_text):
                        log(log_file, f"ActiveX text (first 500 chars): {activex_text[:500]}")
                        window_text = activex_text
                except Exception as e:
                    log(log_file, f"Could not get ActiveX text: {e}", "WARNING")
        
        # Method 3: Try to get text from child elements
        try:
            log(log_file, "Attempting to get text from child elements...")
            children = window.children()
            log(log_file, f"Found {len(children)} child element(s)")
            
            all_text = []
            for i, child in enumerate(children[:20]):  # Check more children
                try:
                    child_text = child.window_text()
                    if child_text and child_text.strip():
                        all_text.append(child_text)
                        log(log_file, f"  Child {i+1} text: {child_text[:100]}...")
                except:
                    pass
            
            if all_text:
                combined_text = " ".join(all_text)
                log(log_file, f"Combined child text (first 500 chars): {combined_text[:500]}")
                if len(combined_text) > len(window_text):
                    window_text = combined_text
        except Exception as e:
            log(log_file, f"Could not get child text: {e}", "WARNING")
        
        # Method 4: Try to get element info
        try:
            log(log_file, "Attempting to get element info...")
            element_info = window.element_info
            log(log_file, f"Element name: {element_info.name}")
            log(log_file, f"Element class: {element_info.class_name}")
            log(log_file, f"Element control type: {element_info.control_type}")
        except Exception as e:
            log(log_file, f"Could not get element info: {e}", "WARNING")
        
        if window_text and len(window_text.strip()) > 0:
            return window_text
        else:
            log(log_file, "No text extracted from window", "WARNING")
            return None
        
    except Exception as e:
        log(log_file, f"Error extracting text: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return None


def find_print_button(window, log_file):
    """Find print button in preview window.
    For ActiveX controls, we need to search within child windows recursively.
    """
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 3: Finding print button")
    log(log_file, "="*60)
    
    if not window:
        log(log_file, "No window provided", "ERROR")
        return None
    
    try:
        log(log_file, "Searching for print button...")
        window_title = window.window_text()
        window_class = window.class_name()
        log(log_file, f"Window: '{window_title}' (Class: {window_class})")
        
        # Method 1: Search by button text in main window
        button_texts = ["Print", "Print...", "Print Label", "Print Preview"]
        for text in button_texts:
            try:
                button = window.child_window(title=text, control_type="Button")
                if button.exists():
                    log(log_file, f"✓ Found print button by text: '{text}'")
                    return button
            except:
                pass
        
        # Method 2: For IEFrame windows or Edge windows, look for IEFrame child and toolbar
        if "IEFrame" in window_class or "Chrome_WidgetWin_1" in window_class:
            log(log_file, f"{window_class} detected - searching for IEFrame child and toolbar...")
            try:
                # Find IEFrame child (might be nested in Edge windows)
                ie_frame = None
                if "IEFrame" in window_class:
                    ie_frame = window
                else:
                    # For Edge windows, look for IEFrame child
                    try:
                        ie_frame = window.child_window(class_name="IEFrame")
                        if ie_frame.exists():
                            log(log_file, "Found IEFrame child in Edge window")
                    except:
                        pass
                
                if ie_frame:
                    # Look for toolbar in IEFrame
                    try:
                        toolbars = ie_frame.descendants(class_name="ToolbarWindow32")
                        log(log_file, f"Found {len(toolbars)} toolbar(s) in IEFrame")
                        for i, toolbar in enumerate(toolbars):
                            try:
                                button_count = toolbar.button_count()
                                log(log_file, f"  Toolbar {i+1}: {button_count} button(s)")
                                for btn_idx in range(button_count):
                                    try:
                                        button_info = toolbar.button(btn_idx)
                                        if button_info:
                                            # Access attributes - text and tooltip are methods, not properties
                                            try:
                                                btn_text = button_info.text() if callable(button_info.text) else (button_info.text if hasattr(button_info, 'text') else '')
                                            except:
                                                btn_text = ''
                                            
                                            try:
                                                btn_tooltip = button_info.tooltip() if callable(button_info.tooltip) else (button_info.tooltip if hasattr(button_info, 'tooltip') else '')
                                            except:
                                                btn_tooltip = ''
                                            
                                            # Also try to get text using toolbar methods
                                            if not btn_text or not isinstance(btn_text, str):
                                                try:
                                                    btn_text = toolbar.button_text(btn_idx)
                                                except:
                                                    pass
                                            
                                            log(log_file, f"    Button {btn_idx}: text='{btn_text}', tooltip='{btn_tooltip}'")
                                            if isinstance(btn_text, str) and isinstance(btn_tooltip, str) and ("print" in btn_text.lower() or "print" in btn_tooltip.lower()):
                                                log(log_file, f"  → Found print button at index {btn_idx} in IEFrame toolbar")
                                                return (toolbar, btn_idx)
                                    except Exception as e:
                                        log(log_file, f"    Button {btn_idx}: Error - {e}", "WARNING")
                            except Exception as e:
                                log(log_file, f"  Toolbar {i+1}: Error - {e}", "WARNING")
                    except Exception as e:
                        log(log_file, f"Could not search toolbars in IEFrame: {e}", "WARNING")
                    
                    # Also try to find Internet Explorer_Server for standard buttons
                    try:
                        ie_server = ie_frame.child_window(class_name="Internet Explorer_Server")
                        if ie_server.exists():
                            log(log_file, "Found Internet Explorer_Server child")
                            try:
                                buttons = ie_server.descendants(control_type="Button")
                                log(log_file, f"Found {len(buttons)} button(s) in ActiveX control")
                                for i, button in enumerate(buttons):
                                    try:
                                        button_text = button.window_text()
                                        log(log_file, f"  Button {i+1}: text='{button_text}'")
                                        if "print" in button_text.lower():
                                            log(log_file, f"  → Found print button: '{button_text}'")
                                            return button
                                    except Exception as e:
                                        log(log_file, f"  Button {i+1}: Error - {e}", "WARNING")
                            except Exception as e:
                                log(log_file, f"Could not search buttons in ActiveX: {e}", "WARNING")
                    except Exception as e:
                        log(log_file, f"Could not find Internet Explorer_Server: {e}", "WARNING")
            except Exception as e:
                log(log_file, f"Error searching IEFrame: {e}", "WARNING")
        
        # Method 3: Search all buttons recursively in window and children
        try:
            log(log_file, "Searching all buttons recursively...")
            buttons = window.descendants(control_type="Button")
            log(log_file, f"Found {len(buttons)} button(s) total (including children)")
            
            for i, button in enumerate(buttons):
                try:
                    button_text = button.window_text()
                    button_class = button.class_name()
                    log(log_file, f"  Button {i+1}: text='{button_text}', class='{button_class}'")
                    
                    if "print" in button_text.lower():
                        log(log_file, f"  → Potential print button found: '{button_text}'")
                        return button
                except Exception as e:
                    log(log_file, f"  Button {i+1}: Error getting info - {e}", "WARNING")
        except Exception as e:
            log(log_file, f"Error searching buttons: {e}", "WARNING")
        
        # Method 4: Search for ToolbarWindow32 controls (toolbar buttons)
        try:
            log(log_file, "Searching for toolbar controls...")
            toolbars = window.descendants(class_name="ToolbarWindow32")
            log(log_file, f"Found {len(toolbars)} toolbar(s)")
            
            # Track the largest toolbar for fallback
            largest_toolbar = None
            largest_button_count = 0
            
            for i, toolbar in enumerate(toolbars):
                try:
                    log(log_file, f"  Toolbar {i+1}:")
                    # Get toolbar button count
                    try:
                        button_count = toolbar.button_count()
                        log(log_file, f"    Button count: {button_count}")
                        
                        # Track largest toolbar for fallback
                        if button_count > largest_button_count:
                            largest_button_count = button_count
                            largest_toolbar = toolbar
                        
                        # Check each button in the toolbar
                        # Track visible buttons separately
                        visible_buttons = []
                        for btn_idx in range(button_count):
                            try:
                                # Get button info (text, tooltip, etc.)
                                # toolbar.button() returns a _toolbar_button object with attributes
                                button_info = toolbar.button(btn_idx)
                                if button_info:
                                    # Check if button is visible using Windows API
                                    is_visible = True
                                    if WIN32_AVAILABLE:
                                        try:
                                            toolbar_handle = toolbar.handle
                                            # TB_GETSTATE = 0x0411 (WM_USER + 17)
                                            # Returns button state flags
                                            TB_GETSTATE = 0x0411
                                            TBSTATE_HIDDEN = 0x0008
                                            state = win32api.SendMessage(toolbar_handle, TB_GETSTATE, btn_idx, 0)
                                            if state & TBSTATE_HIDDEN:
                                                is_visible = False
                                                log(log_file, f"    Button {btn_idx}: Hidden (skipping)")
                                                continue
                                        except:
                                            pass
                                    
                                    # Try to get button rectangle to verify it's actually visible
                                    btn_rect = None
                                    try:
                                        # Try using Windows API to get button rectangle
                                        if WIN32_AVAILABLE:
                                            try:
                                                toolbar_handle = toolbar.handle
                                                # TB_GETRECT = 0x0433 (WM_USER + 51)
                                                TB_GETRECT = 0x0433
                                                rect = win32gui.GetClientRect(toolbar_handle)
                                                # Create RECT structure
                                                import ctypes
                                                class RECT(ctypes.Structure):
                                                    _fields_ = [("left", ctypes.c_int),
                                                               ("top", ctypes.c_int),
                                                               ("right", ctypes.c_int),
                                                               ("bottom", ctypes.c_int)]
                                                button_rect = RECT()
                                                result = win32api.SendMessage(toolbar_handle, TB_GETRECT, btn_idx, ctypes.addressof(button_rect))
                                                if result:
                                                    # Convert to screen coordinates
                                                    toolbar_rect = toolbar.rectangle()
                                                    btn_rect = (toolbar_rect.left + button_rect.left,
                                                               toolbar_rect.top + button_rect.top,
                                                               toolbar_rect.left + button_rect.right,
                                                               toolbar_rect.top + button_rect.bottom)
                                                    # Check if button has valid size (not a separator)
                                                    if button_rect.right - button_rect.left < 5:
                                                        log(log_file, f"    Button {btn_idx}: Separator (skipping)")
                                                        continue
                                            except:
                                                pass
                                    except:
                                        pass
                                    
                                    # Access attributes - text and tooltip are methods, not properties
                                    try:
                                        btn_text = button_info.text() if callable(button_info.text) else (button_info.text if hasattr(button_info, 'text') else '')
                                    except:
                                        btn_text = ''
                                    
                                    try:
                                        btn_tooltip = button_info.tooltip() if callable(button_info.tooltip) else (button_info.tooltip if hasattr(button_info, 'tooltip') else '')
                                    except:
                                        btn_tooltip = ''
                                    
                                    # Also try alternative methods to get button text
                                    if not btn_text or not isinstance(btn_text, str):
                                        try:
                                            # Try toolbar.button_text() method
                                            btn_text = toolbar.button_text(btn_idx)
                                        except:
                                            pass
                                    
                                    # Try to get tooltip using Windows API
                                    if not btn_tooltip and WIN32_AVAILABLE:
                                        try:
                                            toolbar_handle = toolbar.handle
                                            # TB_GETBUTTONTEXT = 0x041B (WM_USER + 27) for text
                                            # For tooltip, we need to use TTN_GETDISPINFO or get it from button info
                                            # Let's try TB_GETBUTTONINFO
                                            TB_GETBUTTONINFO = 0x0441  # WM_USER + 65
                                            # This is complex, so let's skip for now
                                        except:
                                            pass
                                    
                                    if btn_rect:
                                        log(log_file, f"    Button {btn_idx}: rect={btn_rect}, text='{btn_text}', tooltip='{btn_tooltip}'")
                                    else:
                                        log(log_file, f"    Button {btn_idx}: text='{btn_text}', tooltip='{btn_tooltip}'")
                                    
                                    # Track visible buttons
                                    visible_buttons.append((btn_idx, btn_text, btn_tooltip, btn_rect))
                                    
                                    # Check if it's a print button
                                    if "print" in btn_text.lower() or "print" in btn_tooltip.lower():
                                        log(log_file, f"  → Found print button at index {btn_idx} in toolbar {i+1}")
                                        # Return toolbar and button index as a tuple
                                        return (toolbar, btn_idx)
                            except Exception as e:
                                log(log_file, f"    Button {btn_idx}: Error getting info - {e}", "WARNING")
                        
                        # Log visible button count
                        log(log_file, f"    Visible buttons: {len(visible_buttons)} out of {button_count} total")
                        
                        # Only use this toolbar if it has multiple visible buttons (the main toolbar, not zoom toolbar)
                        # The zoom toolbar typically has 1 button, the main toolbar should have 8+ buttons
                        if visible_buttons and len(visible_buttons) >= 3 and i == 1:  # Second toolbar (i=1) with the main toolbar
                            first_visible_idx = visible_buttons[0][0]
                            log(log_file, f"    No print button found by text, trying first visible button (index {first_visible_idx})")
                            log(log_file, f"    Note: Using toolbar with {len(visible_buttons)} visible buttons")
                            return (toolbar, first_visible_idx)
                        elif visible_buttons and len(visible_buttons) < 3:
                            log(log_file, f"    Skipping toolbar {i+1}: Only {len(visible_buttons)} visible button(s) (likely zoom toolbar)")
                    except Exception as e:
                        log(log_file, f"    Could not get button count: {e}", "WARNING")
                    
                except Exception as e:
                    log(log_file, f"  Toolbar {i+1}: Error - {e}", "WARNING")
            
            # Note: We now track visible buttons per toolbar, so no global fallback needed
        except Exception as e:
            log(log_file, f"Error searching toolbars: {e}", "WARNING")
        
        # Method 5: Try to access children directly and look for toolbars
        try:
            log(log_file, "Trying to access window children directly...")
            children = window.children()
            log(log_file, f"Found {len(children)} direct child(ren)")
            for i, child in enumerate(children):
                try:
                    child_title = child.window_text()
                    child_class = child.class_name()
                    log(log_file, f"  Child {i+1}: title='{child_title}', class='{child_class}'")
                    
                    # Check if this child is a toolbar
                    if "ToolbarWindow32" in child_class:
                        try:
                            button_count = child.button_count()
                            log(log_file, f"    Found toolbar with {button_count} button(s)")
                            for btn_idx in range(button_count):
                                try:
                                    button_info = child.button(btn_idx)
                                    if button_info:
                                        # Access attributes directly, not as dict keys
                                        try:
                                            btn_text = button_info.text() if callable(button_info.text) else (button_info.text if hasattr(button_info, 'text') else '')
                                        except:
                                            btn_text = ''
                                        
                                        try:
                                            btn_tooltip = button_info.tooltip() if callable(button_info.tooltip) else (button_info.tooltip if hasattr(button_info, 'tooltip') else '')
                                        except:
                                            btn_tooltip = ''
                                        
                                        # Also try to get text using toolbar methods
                                        if not btn_text or not isinstance(btn_text, str):
                                            try:
                                                btn_text = child.button_text(btn_idx)
                                            except:
                                                pass
                                        
                                        log(log_file, f"      Button {btn_idx}: text='{btn_text}', tooltip='{btn_tooltip}'")
                                        if "print" in btn_text.lower() or "print" in btn_tooltip.lower():
                                            log(log_file, f"  → Found print button at index {btn_idx}")
                                            return (child, btn_idx)
                                except Exception as e:
                                    log(log_file, f"      Button {btn_idx}: Error - {e}", "WARNING")
                        except Exception as e:
                            log(log_file, f"    Could not access toolbar buttons: {e}", "WARNING")
                    
                    # Try to find standard buttons in this child
                    try:
                        child_buttons = child.descendants(control_type="Button")
                        for btn in child_buttons:
                            try:
                                btn_text = btn.window_text()
                                if "print" in btn_text.lower():
                                    log(log_file, f"  → Found print button in child: '{btn_text}'")
                                    return btn
                            except:
                                pass
                    except:
                        pass
                except Exception as e:
                    log(log_file, f"  Child {i+1}: Error - {e}", "WARNING")
        except Exception as e:
            log(log_file, f"Error accessing children: {e}", "WARNING")
        
        log(log_file, "Print button not found by text search", "WARNING")
        log(log_file, "Note: ActiveX controls may require coordinate-based clicking")
        log(log_file, "You may need to use image recognition or manual coordinate clicking")
        log(log_file, "WARNING: Print button not found", "WARNING")
        log(log_file, "You may need to manually identify the button location", "WARNING")
        
        return None
        
    except Exception as e:
        log(log_file, f"Error finding print button: {type(e).__name__}: {str(e)}", "ERROR")
        import traceback
        log(log_file, traceback.format_exc(), "ERROR")
        return None


def click_print_button(button, log_file, preview_window=None):
    """Click the print button.
    Button can be either:
    - A standard button control
    - A tuple of (toolbar, button_index) for toolbar buttons
    
    Args:
        button: Button control or (toolbar, button_index) tuple
        log_file: Log file handle
        preview_window: Optional preview window to focus before clicking
    """
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 4: Clicking print button")
    log(log_file, "="*60)
    
    if not button:
        log(log_file, "No print button provided", "ERROR")
        return False
    
    try:
        # Check if it's a toolbar button (tuple) or regular button
        if isinstance(button, tuple) and len(button) == 2:
            # It's a toolbar button: (toolbar, button_index)
            toolbar, button_index = button
            log(log_file, f"Toolbar button detected (index: {button_index})")
            
            try:
                # Get button info for verification
                button_info = toolbar.button(button_index)
                if button_info:
                    # Access attributes - text and tooltip are methods, not properties
                    try:
                        btn_text = button_info.text() if callable(button_info.text) else (button_info.text if hasattr(button_info, 'text') else '')
                    except:
                        btn_text = ''
                    
                    try:
                        btn_tooltip = button_info.tooltip() if callable(button_info.tooltip) else (button_info.tooltip if hasattr(button_info, 'tooltip') else '')
                    except:
                        btn_tooltip = ''
                    
                    # Also try to get text using toolbar methods
                    if not btn_text or not isinstance(btn_text, str):
                        try:
                            btn_text = toolbar.button_text(button_index)
                        except:
                            pass
                    
                    log(log_file, f"Button info: text='{btn_text}', tooltip='{btn_tooltip}'")
            except Exception as e:
                log(log_file, f"Could not get button info: {e}", "WARNING")
            
            # Check if toolbar is visible
            if not toolbar.is_visible():
                log(log_file, "Toolbar is not visible", "WARNING")
            
            # Verify we're clicking on the correct window (preview window, not search window)
            if preview_window:
                try:
                    # Get preview window title to verify
                    preview_title = preview_window.window_text()
                    log(log_file, f"Verifying preview window: '{preview_title}'")
                    
                    # Check if preview window still exists and is the correct one
                    if not preview_window.exists():
                        log(log_file, "ERROR: Preview window no longer exists!", "ERROR")
                        return False
                    
                    # Focus the preview window first to ensure clicks go to the right window
                    log(log_file, "Focusing preview window first...")
                    preview_window.set_focus()
                    time.sleep(0.3)  # Wait for focus
                    
                    # Verify focus was successful
                    try:
                        focused_window = Desktop(backend="win32").top_window()
                        focused_title = focused_window.window_text()
                        if "enlabel" in focused_title.lower() and "work" in focused_title.lower():
                            log(log_file, f"✓ Preview window focused: '{focused_title}'")
                        else:
                            log(log_file, f"WARNING: Focused window is '{focused_title}' (may not be preview window)", "WARNING")
                    except:
                        pass
                except Exception as e:
                    log(log_file, f"Could not focus/verify preview window: {e}", "ERROR")
                    return False
            
            # Also try to focus the parent window
            try:
                parent_window = toolbar.parent()
                if parent_window:
                    parent_title = parent_window.window_text()
                    log(log_file, f"Toolbar parent window: '{parent_title}'")
                    # Only focus if it's the preview window
                    if preview_window and parent_title == preview_window.window_text():
                        log(log_file, "Focusing toolbar parent window...")
                        parent_window.set_focus()
                        time.sleep(0.2)  # Wait for focus
            except Exception as e:
                log(log_file, f"Could not focus parent window: {e}", "WARNING")
            
            # Verify button index is valid before clicking
            try:
                button_count = toolbar.button_count()
                if button_index >= button_count:
                    log(log_file, f"ERROR: Button index {button_index} is out of range (toolbar has {button_count} buttons)", "ERROR")
                    return False
                
                # Safety check: Don't click if toolbar has only 1 button (likely zoom toolbar)
                if button_count <= 1:
                    log(log_file, f"ERROR: Toolbar has only {button_count} button(s) - likely zoom toolbar, not main toolbar!", "ERROR")
                    log(log_file, "Refusing to click to prevent clicking wrong toolbar", "ERROR")
                    return False
            except Exception as e:
                log(log_file, f"Could not verify button count: {e}", "WARNING")
            
            # Verify we're still on the preview window before clicking
            if preview_window:
                try:
                    if not preview_window.exists():
                        log(log_file, "ERROR: Preview window disappeared before clicking!", "ERROR")
                        return False
                except:
                    pass
            
            log(log_file, f"Clicking toolbar button (index {button_index})...")
            log(log_file, f"WARNING: This will click the button. Ensure preview window is focused!", "WARNING")
            
            # Method 1: Use Windows API to send TB_PRESSBUTTON message directly
            if WIN32_AVAILABLE:
                try:
                    toolbar_handle = toolbar.handle
                    
                    # Verify toolbar handle is valid
                    if not toolbar_handle:
                        log(log_file, "ERROR: Invalid toolbar handle", "ERROR")
                        return False
                    
                    # TB_PRESSBUTTON = 0x0401 (WM_USER + 1)
                    # wParam = button index, lParam = MAKELONG(TRUE, 0) for press, MAKELONG(FALSE, 0) for release
                    TB_PRESSBUTTON = 0x0401
                    
                    # Double-check button index is valid
                    try:
                        # Try to get button state first to verify it exists
                        TB_GETSTATE = 0x0411
                        state = win32api.SendMessage(toolbar_handle, TB_GETSTATE, button_index, 0)
                        log(log_file, f"Button state: {hex(state)}")
                    except Exception as e:
                        log(log_file, f"WARNING: Could not get button state: {e}", "WARNING")
                    
                    # Press button (1 = press, 0 = release)
                    # First press
                    result1 = win32api.SendMessage(toolbar_handle, TB_PRESSBUTTON, button_index, win32api.MAKELONG(1, 0))
                    time.sleep(0.05)
                    # Then release
                    result2 = win32api.SendMessage(toolbar_handle, TB_PRESSBUTTON, button_index, win32api.MAKELONG(0, 0))
                    log(log_file, f"✓ Toolbar button clicked (method 1: Windows API TB_PRESSBUTTON, results: {result1}, {result2})")
                    
                    # Wait and verify preview window still exists (didn't crash or reopen search window)
                    time.sleep(1)
                    if preview_window:
                        try:
                            if preview_window.exists():
                                current_title = preview_window.window_text()
                                log(log_file, f"✓ Preview window still exists after click: '{current_title}'")
                                
                                # Check if the window title changed (might indicate search window reopened)
                                if "search" in current_title.lower() and "preview" not in current_title.lower():
                                    log(log_file, "ERROR: Window title suggests search window reopened instead of print dialog!", "ERROR")
                                    return False
                            else:
                                log(log_file, "WARNING: Preview window closed after click", "WARNING")
                        except Exception as e:
                            log(log_file, f"WARNING: Could not verify preview window after click: {e}", "WARNING")
                    
                    time.sleep(1)  # Additional wait for print dialog
                    return True
                except Exception as e:
                    log(log_file, f"Windows API method failed: {e}", "WARNING")
                    import traceback
                    log(log_file, traceback.format_exc(), "ERROR")
            
            # Method 2: Try to get actual button rectangle and click at its center
            try:
                button_info = toolbar.button(button_index)
                if button_info:
                    try:
                        # Try to get actual button rectangle
                        btn_rect = button_info.rect() if callable(button_info.rect) else (button_info.rect if hasattr(button_info, 'rect') else None)
                        if btn_rect:
                            center_x = btn_rect.left + (btn_rect.width() // 2)
                            center_y = btn_rect.top + (btn_rect.height() // 2)
                            log(log_file, f"Button rectangle: {btn_rect}")
                            log(log_file, f"Button center: ({center_x}, {center_y})")
                            toolbar.click_input(coords=(center_x, center_y))
                            log(log_file, "✓ Toolbar button clicked (method 1: actual button rectangle)")
                            time.sleep(2)
                            return True
                    except Exception as e:
                        log(log_file, f"Could not get button rectangle: {e}", "WARNING")
            except Exception as e:
                log(log_file, f"Method 1 failed: {e}", "WARNING")
            
            # Method 3: Calculate button position based on toolbar layout with better coordinates
            try:
                # Get toolbar rectangle first (screen coordinates)
                toolbar_rect = toolbar.rectangle()
                log(log_file, f"Toolbar rectangle (screen coords): {toolbar_rect}")
                
                # Focus the toolbar window first
                try:
                    toolbar.set_focus()
                    time.sleep(0.2)
                except:
                    pass
                
                # Get button info to estimate position
                # Toolbar buttons are typically arranged horizontally
                # First button is usually at the left, we need to calculate position
                button_count = toolbar.button_count()
                toolbar_width = toolbar_rect.width()
                
                # Toolbar buttons are typically 16-24 pixels wide, with some spacing
                # For the first button, try smaller offsets first (it's usually closer to the left edge)
                if button_index == 0:
                    # First button - try multiple offsets starting from smaller values
                    # The first button is typically very close to the left edge (5-15 pixels)
                    offsets_to_try = [8, 12, 15, 20, 25, 10, 18]
                    for offset in offsets_to_try:
                        try:
                            # Use screen coordinates
                            button_x = toolbar_rect.left + offset
                            button_y = toolbar_rect.top + (toolbar_rect.height() // 2)
                            log(log_file, f"Trying click at offset {offset}: screen coords ({button_x}, {button_y})")
                            
                            # Try using Desktop click_input with screen coordinates
                            desktop = Desktop(backend="win32")
                            desktop.click_input(coords=(button_x, button_y))
                            log(log_file, f"✓ Toolbar button clicked (method 3: Desktop.click_input, offset {offset})")
                            time.sleep(2)
                            return True
                        except Exception as e:
                            log(log_file, f"  Offset {offset} failed: {e}", "WARNING")
                            continue
                else:
                    # Other buttons - estimate based on button width
                    estimated_button_width = 20  # Typical toolbar button width
                    button_x = toolbar_rect.left + (button_index * estimated_button_width) + (estimated_button_width // 2)
                    button_y = toolbar_rect.top + (toolbar_rect.height() // 2)
                    
                    log(log_file, f"Calculated button center (screen coords): ({button_x}, {button_y})")
                    log(log_file, "Clicking at calculated coordinates...")
                    # Use Desktop click_input with screen coordinates
                    desktop = Desktop(backend="win32")
                    desktop.click_input(coords=(button_x, button_y))
                    log(log_file, "✓ Toolbar button clicked (method 3: Desktop.click_input)")
                    time.sleep(2)
                    return True
            except Exception as e:
                log(log_file, f"Method 3 failed: {e}", "WARNING")
                
                # Method 3: Try different coordinate offsets and double-click
                try:
                    toolbar_rect = toolbar.rectangle()
                    # Try multiple offsets for the first button (print button)
                    offsets = [20, 25, 30, 15]  # Try different offsets
                    for offset in offsets:
                        try:
                            click_x = toolbar_rect.left + offset
                            click_y = toolbar_rect.top + (toolbar_rect.height() // 2)
                            log(log_file, f"Trying click at offset {offset}: ({click_x}, {click_y})")
                            # Try double-click for toolbar buttons
                            toolbar.double_click_input(coords=(click_x, click_y))
                            log(log_file, f"✓ Toolbar button double-clicked (method 3: offset {offset})")
                            time.sleep(2)
                            return True
                        except:
                            continue
                    
                    # If double-click didn't work, try single click with last offset
                    click_x = toolbar_rect.left + 25
                    click_y = toolbar_rect.top + (toolbar_rect.height() // 2)
                    toolbar.click_input(coords=(click_x, click_y))
                    log(log_file, f"✓ Toolbar button clicked (method 3: single click)")
                    time.sleep(2)
                    return True
                except Exception as e:
                    log(log_file, f"Method 3 failed: {e}", "WARNING")
                    log(log_file, "All toolbar click methods failed", "ERROR")
                    return False
        else:
            # It's a regular button control
            log(log_file, "Standard button detected")
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
    """Close the preview window with verification."""
    log(log_file, "\n" + "="*60)
    log(log_file, "STEP 5: Closing preview window")
    log(log_file, "="*60)
    
    if not window:
        log(log_file, "No window provided", "ERROR")
        return False
    
    try:
        # Verify window identity before closing
        window_title = window.window_text()
        window_class = window.class_name()
        log(log_file, f"Verifying window to close:")
        log(log_file, f"    Title: '{window_title}'")
        log(log_file, f"    Class: '{window_class}'")
        
        # Safety check: Make sure this is actually a preview window
        title_lower = window_title.lower()
        class_lower = window_class.lower()
        is_preview_window = (
            "enlabel" in title_lower or 
            "internet explorer" in title_lower or
            "ieframe" in class_lower or
            "internet explorer_server" in class_lower
        )
        
        if not is_preview_window:
            log(log_file, "WARNING: Window does not appear to be a preview window!", "WARNING")
            log(log_file, "This might be the wrong window. Aborting close operation.", "WARNING")
            response = input("Are you sure you want to close this window? (yes/no): ").strip().lower()
            if response != "yes":
                log(log_file, "Close operation cancelled by user")
                return False
        
        # Method 1: Try to find and click close button
        log(log_file, "Method 1: Searching for close button...")
        close_texts = ["Close", "X", "Cancel", "Exit"]
        for text in close_texts:
            try:
                close_button = window.child_window(title=text, control_type="Button")
                if close_button.exists():
                    log(log_file, f"Found close button by text: '{text}'")
                    log(log_file, "Clicking close button...")
                    close_button.click_input()
                    log(log_file, "✓ Close button clicked")
                    time.sleep(1)
                    
                    # Verify window is closed
                    try:
                        if not window.exists():
                            log(log_file, "✓ Window successfully closed (verified)")
                            return True
                        else:
                            log(log_file, "Window still exists after close button click", "WARNING")
                    except Exception as verify_error:
                        # Window object may be invalid after close - this is success
                        log(log_file, f"Window object invalid after close - likely closed successfully: {verify_error}", "INFO")
                        log(log_file, "✓ Window successfully closed (object invalidated)")
                        return True
            except Exception as e:
                log(log_file, f"Could not find/click close button '{text}': {e}", "WARNING")
        
        # Method 2: Try to close window directly (for IEFrame windows)
        log(log_file, "Method 2: Attempting to close window directly...")
        try:
            # For IEFrame, we might need to close the parent application
            if "IEFrame" in window_class:
                log(log_file, "IEFrame detected - using close() method")
            window.close()
            log(log_file, "✓ Window.close() called")
            time.sleep(1)
            
            # Verify window is closed
            # Note: After close(), the window object may become invalid, so we need to handle that
            try:
                if not window.exists():
                    log(log_file, "✓ Window successfully closed (verified)")
                    return True
                else:
                    log(log_file, "Window still exists after close()", "WARNING")
            except Exception as verify_error:
                # If exists() fails, it might mean the window was closed and the object is invalid
                # This is actually a success case - the window is gone
                log(log_file, f"Window object invalid after close() - window likely closed successfully: {verify_error}", "INFO")
                log(log_file, "✓ Window successfully closed (object invalidated)")
                return True
        except Exception as e:
            log(log_file, f"Could not close window directly: {e}", "WARNING")
        
        # Method 3: Send Alt+F4 to the window
        log(log_file, "Method 3: Attempting Alt+F4...")
        try:
            # Focus the window first
            window.set_focus()
            time.sleep(0.2)
            window.type_keys("%{F4}")  # Alt+F4
            log(log_file, "✓ Sent Alt+F4")
            time.sleep(1)
            
            # Verify window is closed
            try:
                if not window.exists():
                    log(log_file, "✓ Window successfully closed with Alt+F4 (verified)")
                    return True
                else:
                    log(log_file, "Window still exists after Alt+F4", "WARNING")
            except Exception as verify_error:
                # Window object may be invalid after close - this is success
                log(log_file, f"Window object invalid after Alt+F4 - likely closed successfully: {verify_error}", "INFO")
                log(log_file, "✓ Window successfully closed with Alt+F4 (object invalidated)")
                return True
        except Exception as e:
            log(log_file, f"Could not send Alt+F4: {e}", "WARNING")
        
        log(log_file, "Could not close window with any method", "ERROR")
        log(log_file, "Window may need to be closed manually", "WARNING")
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
                click_print_button(print_button, log_file, preview_window=preview_window)
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
