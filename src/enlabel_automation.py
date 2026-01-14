"""
Enlabel automation for production number search.
Handles login, navigation, and production number search using Selenium.
"""

import time
from pathlib import Path
from typing import Optional
import pandas as pd

from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

from src.logger_setup import get_logger
from src.config_loader import get_config

logger = get_logger(__name__)


class EnlabelAutomation:
    """Automation class for Enlabel website operations."""
    
    def __init__(self, config=None):
        """
        Initialize Enlabel automation.
        
        Args:
            config: Configuration object (optional, will use default if None)
        """
        if config is None:
            config = get_config()
        
        self.config = config
        self.enlabel_config = config.get_section('enlabel')
        self.locators_config = config.get_section('locators')
        self.timeouts_config = config.get_section('timeouts')
        self.paths_config = config.get_section('paths')
        
        self.driver: Optional[webdriver.Edge] = None
        self.wait: Optional[WebDriverWait] = None
        self._filter_initialized = False
        
    def _wait_ready_and_ajax(self, timeout: int = None):
        """
        Wait for document.readyState == 'complete' and for jQuery to be idle (if present).
        
        Args:
            timeout: Timeout in seconds (uses config default if None)
        """
        if timeout is None:
            timeout = self.timeouts_config.get('ajax_wait', 30)
        
        w = WebDriverWait(self.driver, timeout)
        w.until(lambda d: d.execute_script("return document.readyState") == "complete")
        try:
            w.until(lambda d: d.execute_script("return (window.jQuery ? jQuery.active : 0) === 0"))
        except Exception:
            pass  # jQuery not present on all pages
    
    def _switch_into_frame_if_needed(self, locator, probe_timeout: int = 2):
        """
        Ensure Selenium is in the DOM context that contains `locator`.
        Try default content first; otherwise iterate top-level iframes.
        Returns True if found (and switched if needed), else False.
        
        Args:
            locator: Tuple of (By, value) for element locator
            probe_timeout: Timeout for probing frames
        
        Returns:
            True if element found (and context switched if needed), False otherwise
        """
        by, value = locator
        self.driver.switch_to.default_content()
        try:
            WebDriverWait(self.driver, probe_timeout).until(EC.presence_of_element_located(locator))
            return True
        except TimeoutException:
            pass
        
        for fr in self.driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                self.driver.switch_to.default_content()
                self.driver.switch_to.frame(fr)
                WebDriverWait(self.driver, probe_timeout).until(EC.presence_of_element_located(locator))
                return True
            except (TimeoutException, StaleElementReferenceException):
                continue
        
        self.driver.switch_to.default_content()
        return False
    
    def start_browser(self):
        """Initialize browser and WebDriver."""
        logger.info("Starting browser...")
        options = Options()
        
        browser_config = self.config.get_section('browser')
        if browser_config.get('headless', False):
            options.add_argument('--headless')
        
        self.driver = webdriver.Edge(options=options)
        self.wait = WebDriverWait(self.driver, self.timeouts_config.get('element_wait', 10))
        
        logger.info("Browser started successfully")
    
    def login(self):
        """Login to Enlabel website."""
        logger.info("Logging in to Enlabel...")
        
        login_url = self.enlabel_config['login_url']
        username = self.enlabel_config.get('username') or self.config.enlabel_username
        password = self.enlabel_config.get('password') or self.config.enlabel_password
        
        if not username or not password:
            raise ValueError("Enlabel credentials not configured. Set username and password in config.yaml or environment variables.")
        
        login_locators = self.locators_config['login']
        
        # Open login page
        self.driver.get(login_url)
        time.sleep(1)
        
        # Enter username
        username_field = self.wait.until(
            EC.presence_of_element_located((By.ID, login_locators['username_field']))
        )
        username_field.clear()
        username_field.send_keys(username)
        time.sleep(1)
        
        # Enter password
        password_field = self.wait.until(
            EC.presence_of_element_located((By.ID, login_locators['password_field']))
        )
        password_field.clear()
        password_field.send_keys(password)
        time.sleep(1)
        
        # Click login button
        login_button = self.wait.until(
            EC.element_to_be_clickable((By.ID, login_locators['login_button']))
        )
        login_button.click()
        time.sleep(1)
        
        logger.info("Login completed")
    
    def _navigate_to_production_search_pane(self):
        """
        Navigate to production search pane and initialize filters.
        This should be called only once before starting the search loop.
        """
        if self._filter_initialized:
            return
        
        logger.info("Navigating to production search pane...")
        
        prod_search_config = self.locators_config['production_search']
        manage_databases_url = self.enlabel_config['manage_databases_url']
        
        # 1) Go to ManageDatabases
        self.driver.get(manage_databases_url)
        self._wait_ready_and_ajax(self.timeouts_config.get('page_load', 40))
        
        # 2) Find the tables grid (handles iframe if used)
        grid_tables_locator = (By.XPATH, prod_search_config['grid_tables_xpath'])
        if not self._switch_into_frame_if_needed(grid_tables_locator, probe_timeout=3):
            raise TimeoutException("Could not locate the 'gridTables' on ManageDatabases page.")
        
        # 3) Click the link in the second data row
        row2_link_xpath = prod_search_config['row2_link_xpath']
        link_elems = self.driver.find_elements(By.XPATH, row2_link_xpath)
        if not link_elems:
            # Fallback: try first visible row
            row2_link_xpath = "//*[@id[contains(.,'gridTables')]]//tr[contains(@class,'rgRow') or contains(@class,'rgAltRow')][1]//td[1]//a"
            link_elems = self.driver.find_elements(By.XPATH, row2_link_xpath)
        if not link_elems:
            raise TimeoutException("No row link found in gridTables.")
        self.driver.execute_script("arguments[0].click();", link_elems[0])
        
        # 4) Wait for the records view to load
        self._wait_ready_and_ajax(self.timeouts_config.get('page_load', 40))
        self.driver.switch_to.default_content()
        
        # 5) Switch to context with filters/grid
        targets = [
            (By.XPATH, "//*[contains(@id,'gridDbRecords')]"),
            (By.XPATH, prod_search_config['operand_dropdown']),
            (By.XPATH, prod_search_config['column_dropdown']),
            (By.XPATH, "//*[contains(@class,'rgCommandRow') or contains(@class,'rgCommandCell') or contains(@id,'Command')]"),
        ]
        found_context = False
        for loc in targets:
            if self._switch_into_frame_if_needed(loc, probe_timeout=2):
                found_context = True
                break
        if not found_context:
            raise TimeoutException("After opening the record, no grid/filters/command area detected.")
        
        # 6) Click command area to show filters (if needed)
        try:
            cmd = WebDriverWait(self.driver, 4).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(@id,'gridCommand') or contains(@class,'rgCommandRow') or contains(@class,'rgCommandCell') or contains(@id,'Command')]")
                )
            )
            anchors = [a for a in cmd.find_elements(By.XPATH, ".//a[normalize-space()]") if a.is_displayed() and a.is_enabled()]
            if len(anchors) >= 2:
                try:
                    anchors[1].click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", anchors[1])
                self._wait_ready_and_ajax(15)
        except TimeoutException:
            # No command bar; filters may already be visible
            pass
        
        # 7) Set filter dropdowns (one-time setup)
        operand_dd = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, prod_search_config['operand_dropdown']))
        )
        column_dd = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, prod_search_config['column_dropdown']))
        )
        Select(operand_dd).select_by_index(prod_search_config['operand_index'])
        Select(column_dd).select_by_index(prod_search_config['column_index'])
        
        self._filter_initialized = True
        logger.info("Production search pane initialized")
    
    def search_production_number(self, lot_number: str) -> Optional[str]:
        """
        Search for production number using a lot number.
        Assumes the production search pane has already been initialized.
        
        Args:
            lot_number: Lot number to search for
        
        Returns:
            Production number if found, None otherwise
        """
        logger.info(f"Searching for production number with lot: {lot_number}")
        
        prod_search_config = self.locators_config['production_search']
        
        try:
            # Ensure we're in the right context
            self.driver.switch_to.default_content()
            if not self._switch_into_frame_if_needed(
                (By.ID, prod_search_config['value_input']), 
                probe_timeout=3
            ):
                raise TimeoutException("Could not locate filter input field")
            
            # Find and clear the lot input field
            lot_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, prod_search_config['value_input']))
            )
            lot_input.clear()
            lot_input.send_keys(lot_number)
            
            # Click find button
            find_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, prod_search_config['find_button']))
            )
            find_button.click()
            time.sleep(2)
            
            # Extract production number
            production_number_element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, prod_search_config['production_number_xpath']))
            )
            production_number = production_number_element.get_attribute("textContent").strip()
            
            logger.info(f"Found production number: {production_number} for lot: {lot_number}")
            return production_number
            
        except TimeoutException as e:
            logger.error(f"Timeout while searching for lot {lot_number}: {e}")
            return None
        except Exception as e:
            logger.error(f"Error searching for lot {lot_number}: {e}")
            return None
    
    def search_production_numbers(self, items_df: pd.DataFrame) -> pd.DataFrame:
        """
        Search for production numbers for all lot numbers in the DataFrame.
        Skips lot numbers that are already production numbers (9-digit).
        
        Args:
            items_df: DataFrame with columns: item_name, lot, is_production_number
        
        Returns:
            DataFrame with added production_number column
        """
        logger.info(f"Starting production number search for {len(items_df)} items")
        
        # Initialize the search pane (one time)
        self._navigate_to_production_search_pane()
        
        # Create result DataFrame
        result_df = items_df.copy()
        result_df['production_number'] = None
        
        # Loop through items and search
        for idx, row in items_df.iterrows():
            lot_number = str(row['lot']).strip()
            item_name = str(row['item_name']).strip()
            is_production_number = row.get('is_production_number', False)
            
            if is_production_number:
                # Lot number is already a production number (9-digit)
                result_df.at[idx, 'production_number'] = lot_number
                logger.info(f"Lot {lot_number} is already a production number (item: {item_name})")
            else:
                # Search for production number
                production_number = self.search_production_number(lot_number)
                if production_number:
                    result_df.at[idx, 'production_number'] = production_number
                else:
                    logger.warning(f"Could not find production number for lot {lot_number} (item: {item_name})")
        
        logger.info(f"Completed production number search. Found {result_df['production_number'].notna().sum()} production numbers")
        return result_df
    
    def save_production_numbers(self, result_df: pd.DataFrame, filename: str = None):
        """
        Save production numbers to verification CSV file.
        Saves in format: Item number, Lot number, Production number
        
        Args:
            result_df: DataFrame with production numbers (must have item_name, lot, production_number columns)
            filename: Optional filename (defaults to verification CSV)
        """
        # Ensure verification directory exists
        project_root = Path(__file__).parent.parent
        verification_dir = project_root / self.paths_config.get('verification_dir', 'data/verification')
        verification_dir.mkdir(parents=True, exist_ok=True)
        
        if filename is None:
            filename = "production_numbers.csv"
        
        output_file = verification_dir / filename
        
        # Select only the required columns in the correct order
        output_columns = ['item_name', 'lot', 'production_number']
        output_df = result_df[output_columns].copy()
        
        # Rename columns for clarity (Item number, Lot number, Production number)
        output_df.columns = ['Item number', 'Lot number', 'Production number']
        
        # Save to CSV
        output_df.to_csv(output_file, index=False, encoding='utf-8')
        logger.info(f"Saved production numbers to {output_file}")
    
    def close_browser(self):
        """Close browser and cleanup."""
        if self.driver:
            logger.info("Closing browser...")
            self.driver.quit()
            self.driver = None
            self.wait = None
            logger.info("Browser closed")
    
    def __enter__(self):
        """Context manager entry."""
        self.start_browser()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.close_browser()
