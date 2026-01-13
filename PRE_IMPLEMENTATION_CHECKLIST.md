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
  
  - [ ] Search form field selectors
  - [ ] Search button selector
  - [ ] Search results page structure
  - [ ] How are results displayed? (table, list, cards?)
  - [ ] Selector for production number in results
  - [ ] Multiple results handling (first match? filter by item/lot?)
- [ ] **Label Search**:
  - [ ] URL or navigation path from production number to labels
  - [ ] How to access labels? (click link, button, dropdown?)
  - [ ] Selectors for label list/view
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

### Download Process
- [ ] **Download mechanism**:
  - [ ] How are labels downloaded? (download link, download button, context menu?)
  - [ ] Selector for download element
  - [ ] Does download start automatically or require confirmation?
  - [ ] File naming convention (what name does the downloaded file have?)
  - [ ] File format (PDF? Always PDF or variable?)
- [ ] **Browser download behavior**:
  - [ ] Test with Firefox: does it show download dialog?
  - [ ] Default download location
  - [ ] Can downloads be automated via browser preferences?
- [ ] **Download timing**:
  - [ ] How long does download take? (network-dependent, but get ballpark)
  - [ ] Any loading indicators to wait for?

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
- [ ] **Headless mode**:
  - [ ] Can run headless? (may be required if no display)
  - [ ] Does headless work with downloads?
- [ ] **Download configuration**:
  - [ ] Test Firefox download preferences (auto-download to folder)
  - [ ] Document working configuration

### PDF Processing
- [ ] **PDF library capabilities**:
  - [ ] Test PyPDF2 on sample PDFs (if possible)
  - [ ] Test pdfplumber on sample PDFs
  - [ ] Compare extraction quality
  - [ ] Test PDF merging capabilities
- [ ] **OCR setup** (if needed):
  - [ ] Can Tesseract be installed at work?
  - [ ] Test OCR accuracy on sample labels

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
- [ ] **Document work environment details**:
  - [ ] OS version (Windows 10/11?)
  - [ ] Python version
  - [ ] Firefox version
  - [ ] Any other relevant software versions

## 10. Additional Considerations

### Backup Plans
- [ ] **Alternative approaches to research**:
  - [ ] If Selenium doesn't work, research alternative automation tools
  - [ ] If PDF text extraction fails, document OCR setup requirements
  - [ ] If download automation fails, document manual download process

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

---

## Notes Section

_Use this section to write down any additional observations, questions, or findings during your research:_
