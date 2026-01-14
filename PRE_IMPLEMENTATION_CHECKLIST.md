# Pre-Implementation Research Checklist

This document outlines all items that should be investigated and documented at work before starting implementation at home.

## 1. Input Data Structure & Format

### TSV File Analysis

- [x] **Obtain sample TSV files** (at least 2-3 different files if possible)
- [x] **Identify column names** for:
  - [x] Item numbers  
  - [x] Lot numbers 
- [x] **Check data format**:
  - [x] Are item numbers alphanumeric? - Yes
  - [x] Are lot numbers consistent format? - No lot number formats vary.
  - [x] Any special characters or encoding issues? - No
- [x] **Edge cases to document**:
  - [x] Empty/missing values in columns - There can be empty values, but then the item needs to be flagged as manual confirmation will be needed
  - [x] Multiple lots per item - Yes, one item can have multiple lots
  - [x] Special characters in data - No special characters.
  - [x] Date formats if present - No need for date formats

### Invoice PDF Structure

- [x] **Document invoice structure**:
  - [x] Page count (single page or multi-page?) - Multi page
  - [x] Layout/orientation (portrait/landscape) - portrait
  - [x] Where should labels be inserted? (before invoice, after, specific page?) - after the invoice

## 2. Enlabel Website Investigation

### Authentication & Access

- [x] **Login process documentation**:
  - [x] URL of login page - https://pallprod.enlabel.com/Login.aspx?ReturnUrl=%2fDefault.aspx
  - [x] Login form field IDs/names/selectors username ID - ctl00_ContentPlaceHolder1_txtUserName | password field ID - ctl00_ContentPlaceHolder1_txtPassword
  - [x] Submit button selector - ID - ctl00_ContentPlaceHolder1_btnLogin
  - [x] Session timeout duration - not clear at least 15 minutes
  - [x] Any IP restrictions or VPN requirements? needs to be cytiva network, no details on this one, eitherway this will only run from a cytiva computer.
- [x] **Test existing login script** (if available):
  - [x] Does it still work? yes
  - [x] Any changes to the login flow? no
  - [x] Save/note any selector changes no

### Navigation & Search Flow

- [ ] **Production Number Search**:
  - [ ] Exact URL or navigation path to search page. https://pallprod.enlabel.com/Collaboration/ManageDatabases/ManageDatabases.aspx -> #ctl00_MainContent_gridTables_ctl00__1 > td:nth-child(1) > a -> #ctl00_MainContent_gridDbRecords_ctl00_ctl02_ctl00_gridCommand > table > tbody > tr > td:nth-child(2) > a -> #ctl00_MainContent_FilterControl_ddlOperand1 -> #ctl00_MainContent_FilterControl_ddlOperand1 > option:nth-child(2) -> #ctl00_MainContent_FilterControl_ddlColumn1 > option:nth-child(9)
  
- [ ] **Label Search**:
  - [ ] Label search URL - https://pallprod.enlabel.com/ProductionPrint/PrintTypes/PrintByOrder/PrintStart.aspx?ServiceId=10
  - [ ] How to access labels? - after navigating to the label search URL, must input production order number to production order number input field, click next and observe the results.
    Full search flow -> //*[@id="ctl00_MainContent__txtORDER_NUMBER"] -> enter the production order number -> click next (//*[@id="btnNext2"]) -> at this point there is a table element with multiple rows, first row is the header, after that multiple results, need to find the label that fits the requirements, must have Item number, lot number and EPA number.
    table row xpaths -> 1st row - //*[@id="ctl00_MainContent_gridLabels_ctl00__0"] -> 2nd row - //*[@id="ctl00_MainContent_gridLabels_ctl00__1"] -> 3rd row - //*[@id="ctl00_MainContent_gridLabels_ctl00__2"] and so on. as a human The only way to check if what label is it, must preview it and visually confirm if it meets all the requirements - preview button of the first row - //*[@id="ctl00_MainContent_gridLabels_ctl00_ctl04_btnPreview"] -> preview button of the 2nd row - //*[@id="ctl00_MainContent_gridLabels_ctl00_ctl06_btnPreview"] -> preview button 3rd row - //*[@id="ctl00_MainContent_gridLabels_ctl00_ctl08_btnPreview"]
    Here the flow brakes, the preview must be opened from IE mode on microsoft edge, activex is needed to load the label, need to test at work if this is the case. 
- [ ] **Label Identification** (CRITICAL):
  - [ ] Take screenshots of label list/page
  - [ ] How are labels displayed? (list, table, thumbnails?)
  - [ ] What information is visible in the list? (label name, item number, lot number, EPA number, dates?)
  - [ ] DOM structure analysis:
    - [ ] Open browser DevTools (F12)
    - [ ] Inspect label elements
    - [ ] Document HTML structure (copy HTML snippets)
    - [ ] Document CSS classes/IDs used
    - [ ] Document data attributes if present
  - [ ] How to identify the "correct" label?
    - [ ] Match by item number?
    - [ ] Match by lot number?
    - [ ] Match by EPA number?
    - [ ] Match by date (most recent)?
    - [ ] Combination of factors?
  - [ ] What happens if multiple labels match?
  - [ ] What happens if no label matches?
- [ ] **Multiple Windows/Tabs**:
  - [ ] Does clicking labels open new windows/tabs?
  - [ ] Test manually: click a label, see what happens
  - [ ] Document window behavior
  - [ ] Create test script to:
Manually trigger a label preview
Automate clicking print button
Automate printer selection
Automate file save
Integrate into existing EnlabelAutomation class
Add error handling and retry logic
Test with multiple labels

### Download Process
- [ ] **Download mechanism**:
  - [ ] How are labels downloaded? Label data is sent from the server to the client, together with a label template, on the client the label gets constructed and a preview is shown with activex, from the preview window there is a a print button, where the printer gets selected, as we only need the digital copy, Microsoft print to pdf is selected. After that windows file selector dialog pops up to select the saving location and name of the file. after pressing next file is saved.
  - [ ] Selector for download element
  - [ ] Does download start automatically or require confirmation?
  - [ ] File naming convention (what name does the downloaded file have?)
  - [ ] File format (PDF? Always PDF or variable?)
- [ ] **ActiveX Preview Window** (CRITICAL for automation):
  - [ ] Does preview open in new window or same tab?
  - [ ] Window title of preview window (exact text or pattern)
  - [ ] Window class name (use Spy++ or similar tool to inspect)
  - [ ] Screenshot of preview window with print button visible
  - [ ] Print button location:
    - [ ] Is print button always in same position? (relative to window)
    - [ ] Print button text/label (exact text)
    - [ ] Print button appearance (icon, text, color, size)
    - [ ] Can print button be found by image recognition? (take screenshot of button)
  - [ ] How long does it take for preview window to fully load?
  - [ ] Any loading indicators in preview window?
  - [ ] Does preview window have a specific handle/identifier?
  - [ ] Test: Can preview window be found using pywinauto by title/class?
- [ ] **Printer Selection Dialog** (CRITICAL for automation):
  - [ ] Dialog title (exact text, e.g., "Print", "Select Printer", etc.)
  - [ ] Dialog class name (use Spy++ or similar tool)
  - [ ] Screenshot of printer selection dialog
  - [ ] How is "Microsoft Print to PDF" displayed?
    - [ ] In a dropdown list?
    - [ ] In a list box?
    - [ ] Exact text shown (may include version numbers)
  - [ ] How to select "Microsoft Print to PDF"?
    - [ ] Click on it?
    - [ ] Type to search/filter?
    - [ ] Arrow keys to navigate?
  - [ ] Button labels (OK, Print, Next, etc.) - exact text
  - [ ] Button positions or IDs
  - [ ] Does dialog appear immediately after clicking print, or is there a delay?
  - [ ] Test: Can dialog be found using pywinauto?
  - [ ] Test: Can "Microsoft Print to PDF" be selected programmatically?
- [ ] **File Save Dialog** (CRITICAL for automation):
  - [ ] Dialog title (exact text, e.g., "Save Print Output As", "Save As", etc.)
  - [ ] Dialog class name (use Spy++ or similar tool)
  - [ ] Screenshot of file save dialog
  - [ ] File name input field:
    - [ ] Field label/ID
    - [ ] Default filename (if any)
    - [ ] Can field be accessed by pywinauto?
  - [ ] Save location field:
    - [ ] Default save location
    - [ ] Can path be typed directly?
    - [ ] Or must navigate through folder browser?
  - [ ] Save button:
    - [ ] Button text (exact: "Save", "OK", "Next", etc.)
    - [ ] Button position/ID
  - [ ] Cancel button (if needed for error handling)
  - [ ] Does dialog appear immediately after selecting printer, or is there a delay?
  - [ ] Test: Can dialog be found and controlled using pywinauto?
  - [ ] Test: Can file path be entered programmatically?
- [ ] **Browser download behavior**:
  - [ ] Test with Firefox: does it show download dialog?
  - [ ] Default download location
  - [ ] Can downloads be automated via browser preferences?
- [ ] **Download timing**:
  - [ ] How long does download take? (network-dependent, but get ballpark)
  - [ ] Any loading indicators to wait for?
  - [ ] Time from clicking preview to preview window appearing
  - [ ] Time from clicking print to printer dialog appearing
  - [ ] Time from selecting printer to save dialog appearing
  - [ ] Time from clicking save to file actually being saved

### Performance & Timing
- [ ] **Page load times** (rough estimates):
  - [ ] Login page load
  - [ ] Search results load
  - [ ] Label page load
  - [ ] Download initiation
- [ ] **Wait conditions to identify**:
  - [ ] Elements that appear when page is ready
  - [ ] Loading spinners/indicators and their selectors
  - [ ] Elements that disappear when loading completes

### Error Scenarios
- [ ] **Document common errors**:
  - [ ] Invalid login credentials (what error message/element appears?)
  - [ ] Production number not found (what happens?)
  - [ ] No labels found (what is displayed?)
  - [ ] Session timeout (what happens? redirect to login?)
  - [ ] Network errors
- [ ] **Take screenshots** of error states for reference

## 3. Label PDF Structure

### Label Content Analysis
- [ ] **Obtain sample label PDFs** (download 3-5 different labels manually)
- [ ] **Test text extraction** on sample labels:
  - [ ] Try PyPDF2 extraction (quick test script if possible)
  - [ ] Try pdfplumber extraction
  - [ ] Are labels text-based or image-based?
  - [ ] Can item numbers be extracted reliably?
  - [ ] Can lot numbers be extracted reliably?
  - [ ] Can EPA numbers be extracted reliably?
- [ ] **Document label layout**:
  - [ ] Where is item number located? (header, body, footer?)
  - [ ] Where is lot number located?
  - [ ] Where is EPA number located?
  - [ ] Are these in consistent locations across labels?
- [ ] **OCR requirements**:
  - [ ] If text extraction fails, test OCR on sample labels
  - [ ] Take note of label quality (clear text or blurry?)
  - [ ] Estimate OCR accuracy

### Label Format
- [ ] **Page properties**:
  - [ ] Page size (letter, A4, custom?)
  - [ ] Orientation (portrait/landscape)
  - [ ] Single page or multi-page labels?
- [ ] **File properties**:
  - [ ] Typical file size
  - [ ] PDF version

## 4. Work Computer Environment

### Software Installation Restrictions
- [ ] **Python installation**:
  - [ ] Is Python already installed?
  - [ ] What version? (`python --version`)
  - [ ] Can you install Python if not available?
  - [ ] Admin rights required for installation?
- [ ] **Browser availability**:
  - [ ] Is Firefox installed?
  - [ ] What version? (need for Selenium compatibility)
  - [ ] Can you install/update Firefox?
  - [ ] Is geckodriver available? (for Selenium)
- [ ] **Other dependencies**:
  - [ ] Can you install pip packages? (may need admin rights)
  - [ ] Are there firewall/proxy restrictions?
  - [ ] Any software whitelist/blacklist?

### Network & Security
- [ ] **Network access**:
  - [ ] Can access enlabel from work network?
  - [ ] Any VPN required?
  - [ ] Proxy configuration needed?
  - [ ] Firewall rules for outbound connections?
- [ ] **Security policies**:
  - [ ] Can run Python scripts?
  - [ ] Any antivirus that might block Selenium/browser automation?
  - [ ] Credential storage restrictions?
  - [ ] File system access restrictions?

### File System Access
- [ ] **Directory permissions**:
  - [ ] Where can you create/write files? (user directory, network drive?)
  - [ ] Can create project directory structure?
  - [ ] Can write to Downloads folder?
- [ ] **File paths**:
  - [ ] Document typical locations for:
    - [ ] TSV files (source location)
    - [ ] Invoice PDFs (source location)
    - [ ] Where output should be saved
  - [ ] Use absolute paths or relative paths?

### Running Automation
- [ ] **Testing restrictions**:
  - [ ] Can run browser automation during work hours?
  - [ ] Any policies against automation tools?
  - [ ] Impact on work network/performance?

## 5. Business Logic & Requirements

### Workflow Validation
- [ ] **Current manual process**:
  - [ ] Document exact steps currently performed manually
  - [ ] How many items typically processed per run?
  - [ ] How long does manual process take?
- [ ] **Expected output**:
  - [ ] Final PDF structure: labels first then invoice? Or invoice first then labels?
  - [ ] Label order in final PDF (same as TSV order? alphabetical? by production number?)
  - [ ] What if some labels are missing/can't be downloaded?
  - [ ] What if label verification fails?

### Data Validation Rules
- [ ] **Production number mapping**:
  - [ ] Is production number always found for item/lot combinations?
  - [ ] What to do if production number not found?
  - [ ] One-to-one mapping or one-to-many?
- [ ] **Label matching rules**:
  - [ ] Exact match required for item/lot/EPA?
  - [ ] Partial match acceptable?
  - [ ] Case sensitivity?
  - [ ] Whitespace handling?

### Error Handling Requirements
- [ ] **What should happen when**:
  - [ ] Login fails
  - [ ] Production number not found
  - [ ] Label not found
  - [ ] Label download fails
  - [ ] Label verification fails (missing item/lot/EPA)
  - [ ] PDF merge fails
- [ ] **Continue or stop**:
  - [ ] Continue processing other items if one fails?
  - [ ] Generate partial output or require all labels?
  - [ ] Log errors and continue, or fail fast?

## 6. Technical Specifications

### Browser Automation
- [ ] **Selenium compatibility**:
  - [ ] Firefox version at work (need compatible geckodriver)
  - [ ] Test if Selenium can control Firefox (basic test script)
  - [ ] Any browser extensions that interfere?
- [ ] **Edge IE Mode** (CRITICAL for ActiveX):
  - [ ] Is Edge IE mode required for ActiveX to work?
  - [ ] How to enable IE mode in Edge?
  - [ ] Can Selenium control Edge in IE mode?
  - [ ] Test: Open label preview in Edge IE mode manually
  - [ ] Test: Does ActiveX load correctly in Edge IE mode?
  - [ ] Document IE mode configuration steps
  - [ ] Can IE mode be set programmatically via Selenium?
  - [ ] Any compatibility issues with Selenium + Edge IE mode?
- [ ] **Headless mode**:
  - [ ] Can run headless? (may be required if no display)
  - [ ] Does headless work with downloads?
  - [ ] **IMPORTANT**: Headless mode likely won't work with ActiveX/Windows dialogs - document this limitation
- [ ] **Download configuration**:
  - [ ] Test Firefox download preferences (auto-download to folder)
  - [ ] Document working configuration
- [ ] **Multiple Windows Management**:
  - [ ] When preview opens, does it open in new window or tab?
  - [ ] How many windows are open during label printing process?
  - [ ] Can Selenium track all open windows?
  - [ ] Test: Switch between browser window and preview window using Selenium
  - [ ] Test: Can Selenium detect when preview window closes?

### PDF Processing
- [ ] **PDF library capabilities**:
  - [ ] Test PyPDF2 on sample PDFs (if possible)
  - [ ] Test pdfplumber on sample PDFs
  - [ ] Compare extraction quality
  - [ ] Test PDF merging capabilities
- [ ] **OCR setup** (if needed):
  - [ ] Can Tesseract be installed at work?
  - [ ] Test OCR accuracy on sample labels

### Windows Automation Tools
- [ ] **pywinauto installation & testing**:
  - [ ] Can pywinauto be installed? (`pip install pywinauto`)
  - [ ] Test basic pywinauto functionality (find window, click button)
  - [ ] Test finding preview window by title/class
  - [ ] Test finding printer dialog by title/class
  - [ ] Test finding save dialog by title/class
  - [ ] Document any installation issues or restrictions
- [ ] **pyautogui installation & testing**:
  - [ ] Can pyautogui be installed? (`pip install pyautogui`)
  - [ ] Test basic pyautogui functionality (screenshot, click)
  - [ ] Test clicking print button using image recognition
  - [ ] Test coordinate-based clicking (if needed)
  - [ ] Document screen resolution and scaling settings (affects coordinates)
- [ ] **Windows API access**:
  - [ ] Can Python access Windows API? (for advanced automation if needed)
  - [ ] Any restrictions on using ctypes or win32api?
- [ ] **Dialog automation testing**:
  - [ ] Create simple test script to automate one label print manually
  - [ ] Test: Find preview window → Click print → Select printer → Save file
  - [ ] Document any issues or workarounds needed
  - [ ] Test timing: How long to wait between each step?
  - [ ] Test error handling: What if dialog doesn't appear?

### Credentials & Configuration
- [ ] **Authentication method**:
  - [ ] Username/password? (document format requirements)
  - [ ] Any special characters in credentials?
  - [ ] Password change frequency?
  - [ ] Can credentials be stored securely? (environment variables, config file with encryption?)
- [ ] **Configuration needs**:
  - [ ] Document all configurable values:
    - [ ] Timeouts (login, search, download)
    - [ ] Retry counts
    - [ ] Wait intervals
    - [ ] Download folder paths
    - [ ] Output folder paths

## 7. Testing & Validation

### Sample Data Collection
- [ ] **Create test dataset**:
  - [ ] Small TSV file (3-5 items) for initial testing
  - [ ] Corresponding invoice PDF
  - [ ] Known production numbers for test items
  - [ ] Known label characteristics for verification
- [ ] **Document expected results**:
  - [ ] What the final PDF should look like
  - [ ] Expected processing time (rough estimate)
  - [ ] Expected file sizes

### Edge Cases to Test
- [ ] **Prepare edge case examples**:
  - [ ] TSV with missing data
  - [ ] Items with multiple labels
  - [ ] Items with no labels
  - [ ] Very long item/lot numbers
  - [ ] Special characters in item/lot numbers
  - [ ] Large TSV file (many items)
- [ ] **Label printing edge cases**:
  - [ ] What if preview window doesn't open?
  - [ ] What if print button is not visible/clickable?
  - [ ] What if "Microsoft Print to PDF" is not in printer list?
  - [ ] What if file save dialog doesn't appear?
  - [ ] What if file path is too long?
  - [ ] What if file already exists at save location?
  - [ ] What if user has multiple monitors? (dialog might appear on different screen)
  - [ ] What if another application has focus when dialog appears?
  - [ ] What if printer dialog is minimized or behind other windows?

## 8. Documentation to Collect

### Screenshots & Recordings
- [ ] **Screenshot collection**:
  - [ ] Login page
  - [ ] Search page
  - [ ] Search results page
  - [ ] Label list/page
  - [ ] Label detail/view page
  - [ ] Error pages/messages
  - [ ] Download dialog (if appears)
  - [ ] **ActiveX preview window** (full window, with print button visible)
  - [ ] **Print button close-up** (for image recognition if needed)
  - [ ] **Printer selection dialog** (full dialog, showing "Microsoft Print to PDF")
  - [ ] **File save dialog** (full dialog, showing all fields and buttons)
- [ ] **Window inspection data**:
  - [ ] Use Spy++ (or similar tool) to capture:
    - [ ] Preview window: Title, Class Name, Handle
    - [ ] Printer dialog: Title, Class Name, Handle
    - [ ] Save dialog: Title, Class Name, Handle
  - [ ] Document all window properties needed for automation
- [ ] **DOM snapshots**:
  - [ ] Save HTML snippets of key elements (production number results, label list, etc.)
  - [ ] Copy CSS selectors that work
- [ ] **Video recording** (optional but helpful):
  - [ ] Record manual process end-to-end
  - [ ] Helps understand timing and flow

### Notes & Observations
  - [ ] Any quirks or odd behaviors noticed
  - [ ] Timing observations
  - [ ] Common issues encountered manually
  - [ ] Workarounds currently used

## 9. Development Environment Preparation

### Tools & Resources Needed
- [ ] **List software to install at home**:
  - [ ] Python version (match work version if possible)
  - [ ] Firefox (same version as work if possible)
  - [ ] IDE/editor (VS Code, PyCharm, etc.)
  - [ ] Git (for version control)
  - [ ] **Windows inspection tools**:
    - [ ] Spy++ (Windows SDK) or alternative (WinSpy, Window Detective)
    - [ ] For inspecting window titles, class names, handles
- [ ] **Document work environment details**:
  - [ ] OS version (Windows 10/11?)
  - [ ] Python version
  - [ ] Firefox version
  - [ ] Edge version
  - [ ] Screen resolution and scaling (important for pyautogui coordinates)
  - [ ] Number of monitors (affects window positioning)
  - [ ] Any other relevant software versions

## 10. Additional Considerations

### Backup Plans
- [ ] **Alternative approaches to research**:
  - [ ] If Selenium doesn't work, research alternative automation tools
  - [ ] If PDF text extraction fails, document OCR setup requirements
  - [ ] If download automation fails, document manual download process
  - [ ] **If Windows dialog automation fails**:
    - [ ] Can browser print API be used instead? (unlikely with ActiveX)
    - [ ] Can label data be intercepted/downloaded before ActiveX rendering?
    - [ ] Is there an API endpoint that returns label PDF directly?
    - [ ] Can Edge DevTools Protocol be used to intercept print commands?
  - [ ] **If ActiveX doesn't work in automation**:
    - [ ] Can labels be rendered differently? (check browser settings)
    - [ ] Is there a non-ActiveX preview option?
    - [ ] Can labels be accessed via direct download link?

### Compliance & Approval
- [ ] **Check with IT/management** (if required):
  - [ ] Approval for automation tools?
  - [ ] Data handling policies?
  - [ ] Credential storage policies?
  - [ ] Script execution policies?

### Future Enhancements (Document for Later)
- [ ] **Potential improvements** (not for initial version but good to note):
  - [ ] Batch processing multiple TSV files
  - [ ] Email notifications
  - [ ] Integration with other systems
  - [ ] Automated scheduling

---

## Summary: Critical Items (Must-Have Before Starting)

1. ✅ Sample TSV files with column structure documented
2. ✅ Sample invoice PDFs
3. ✅ Sample label PDFs (downloaded manually)
4. ✅ Enlabel DOM structure analysis (selectors, IDs, classes)
5. ✅ Label identification logic (how to find correct label)
6. ✅ Download mechanism understanding
7. ✅ Work computer environment assessment (Python, Firefox, permissions)
8. ✅ Credentials and authentication flow documented
9. ✅ Error scenarios documented with screenshots
10. ✅ PDF text extraction test results
11. ⚠️ **ActiveX Preview Window Details** (window title, class, print button location)
12. ⚠️ **Printer Dialog Details** (title, class, how to select "Microsoft Print to PDF")
13. ⚠️ **File Save Dialog Details** (title, class, field IDs, button labels)
14. ⚠️ **Windows Automation Tools Testing** (pywinauto, pyautogui installation and basic tests)
15. ⚠️ **Edge IE Mode Configuration** (if required for ActiveX)
16. ⚠️ **Window Inspection Data** (Spy++ output for all dialogs)
17. ⚠️ **Timing Information** (how long each dialog takes to appear)

---

## Notes Section

_Use this section to write down any additional observations, questions, or findings during your research:_

### Label Printing Automation Notes

**Key Automation Challenges:**
- ActiveX preview window is not accessible via Selenium (Windows-specific control)
- Must use Windows automation tools (pywinauto/pyautogui) to interact with dialogs
- Headless mode will NOT work - requires visible browser and dialogs
- Multiple windows must be tracked and managed

**Recommended Testing Sequence:**
1. Manually print one label and document every step
2. Use Spy++ to capture window properties (title, class, handle)
3. Test pywinauto to find each window/dialog
4. Test clicking buttons and entering text programmatically
5. Create simple test script to automate one complete label print
6. Test with multiple labels to identify timing issues

**Critical Information to Collect:**
- Exact window titles (may vary by Windows version/language)
- Window class names (more reliable than titles)
- Button positions or IDs
- Timing between each step
- Error scenarios (what if dialog doesn't appear?)
