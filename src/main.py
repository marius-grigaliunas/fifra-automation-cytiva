"""
Main orchestrator for FIFRA Automation.
Coordinates all components and handles the main workflow.
"""

import sys
from pathlib import Path
from typing import Optional

from src.logger_setup import setup_logging, get_logger
from src.config_loader import get_config
from src.data_parser import TSVParser
from src.gui import FIFRAGUI
from src.enlabel_automation import EnlabelAutomation


class FIFRAAutomation:
    """Main automation orchestrator."""
    
    def __init__(self, gui: Optional[FIFRAGUI] = None):
        """
        Initialize the automation orchestrator.
        
        Args:
            gui: GUI instance (optional, will create if None)
        """
        # Setup logging first
        config = get_config()
        logging_config = config.get_section('logging')
        setup_logging(
            log_level=logging_config.get('level', 'INFO'),
            log_file=logging_config.get('file'),
            log_format=logging_config.get('format'),
            max_bytes=logging_config.get('max_bytes', 10485760),
            backup_count=logging_config.get('backup_count', 5)
        )
        
        self.config = config
        self.gui = gui
        self.parser = TSVParser(config)
        
        # Get logger after setup (setup_logging configures root logger, so this works)
        logger = get_logger(__name__)
        logger.info("FIFRA Automation initialized")
    
    def process_files(self, tsv_path: str, invoice_path: str):
        """
        Process TSV and invoice files.
        
        Args:
            tsv_path: Path to TSV file
            invoice_path: Path to invoice PDF file
        """
        try:
            if self.gui:
                self.gui.update_status("Starting file processing...")
                self.gui.update_progress(10)
            
            logger = get_logger(__name__)
            logger.info(f"Processing TSV file: {tsv_path}")
            logger.info(f"Invoice PDF: {invoice_path}")
            
            # Parse TSV file
            if self.gui:
                self.gui.update_status("Parsing TSV file...")
                self.gui.update_progress(20)
            
            parse_result = self.parser.parse_file(tsv_path)
            
            # Display results
            items_df = parse_result['items']
            trip_number = parse_result['trip_number']
            tracking_number = parse_result['tracking_number']
            flagged_rows = parse_result['flagged_rows']
            total_rows = parse_result['total_rows']
            
            # Save parsed data to data/input/parsedInput.tsv
            if self.gui:
                self.gui.update_status("Saving parsed data...")
                self.gui.update_progress(30)
            
            self._save_parsed_data(items_df, trip_number, tracking_number)
            
            # Phase 2.1: Search for production numbers
            if self.gui:
                self.gui.update_status("Searching for production numbers...")
                self.gui.update_progress(50)
            
            production_numbers_df = self._search_production_numbers(items_df)
            
            # Save production numbers to verification file
            if self.gui:
                self.gui.update_status("Saving production numbers...")
                self.gui.update_progress(70)
            
            self._save_production_numbers(production_numbers_df, trip_number, tracking_number)
            
            # Display results in GUI
            if self.gui:
                self.gui.update_status("=" * 60)
                self.gui.update_status("Parsing Complete!")
                self.gui.update_status("=" * 60)
                self.gui.update_status(f"Total rows in TSV: {total_rows}")
                self.gui.update_status(f"Unique item/lot combinations: {len(items_df)}")
                self.gui.update_status(f"Trip Number: {trip_number or 'Not found'}")
                self.gui.update_status(f"Tracking Number: {tracking_number or 'Not found'}")
                
                if flagged_rows:
                    self.gui.update_status(f"\nFlagged rows (need manual confirmation): {len(flagged_rows)}")
                    for flagged in flagged_rows[:5]:  # Show first 5
                        self.gui.update_status(f"  - Row {flagged['index']}: {flagged['issues']}")
                    if len(flagged_rows) > 5:
                        self.gui.update_status(f"  ... and {len(flagged_rows) - 5} more")
                
                self.gui.update_status(f"\nItems to process:")
                for idx, row in items_df.head(10).iterrows():  # Show first 10
                    is_prod = " (is production number)" if row.get('is_production_number') else ""
                    self.gui.update_status(f"  - Item: {row['item_name']}, Lot: {row['lot']}{is_prod}")
                if len(items_df) > 10:
                    self.gui.update_status(f"  ... and {len(items_df) - 10} more items")
                
                self.gui.update_status(f"\nProduction numbers found:")
                found_count = production_numbers_df['production_number'].notna().sum()
                for idx, row in production_numbers_df.head(10).iterrows():  # Show first 10
                    prod_num = row.get('production_number', 'Not found')
                    self.gui.update_status(f"  - Item: {row['item_name']}, Lot: {row['lot']} -> Production: {prod_num}")
                if len(production_numbers_df) > 10:
                    self.gui.update_status(f"  ... and {len(production_numbers_df) - 10} more items")
                
                self.gui.update_status(f"\nParsed data saved to: data/input/parsedInput.tsv")
                self.gui.update_status(f"Production numbers saved to: data/verification/production_numbers.csv")
                self.gui.update_progress(100)
                self.gui.show_completion_message(
                    True,
                    f"Successfully processed {len(items_df)} unique item/lot combinations.\n"
                    f"Found {found_count} production numbers.\n"
                    f"Data saved to data/input/parsedInput.tsv"
                )
            else:
                # Command-line output
                print("=" * 60)
                print("Parsing Complete!")
                print("=" * 60)
                print(f"Total rows in TSV: {total_rows}")
                print(f"Unique item/lot combinations: {len(items_df)}")
                print(f"Trip Number: {trip_number}")
                print(f"Tracking Number: {tracking_number}")
                if flagged_rows:
                    print(f"\nFlagged rows: {len(flagged_rows)}")
                print(f"\nParsed data saved to: data/input/parsedInput.tsv")
                found_count = production_numbers_df['production_number'].notna().sum()
                print(f"Production numbers found: {found_count}/{len(production_numbers_df)}")
                print(f"Production numbers saved to: data/verification/production_numbers.csv")
            
            logger = get_logger(__name__)
            found_count = production_numbers_df['production_number'].notna().sum()
            logger.info(f"Processing complete. {len(items_df)} unique items extracted. {found_count} production numbers found.")
            
        except Exception as e:
            logger = get_logger(__name__)
            error_msg = f"Error processing files: {str(e)}"
            logger.error(error_msg, exc_info=True)
            if self.gui:
                self.gui.update_status(f"ERROR: {error_msg}")
                self.gui.show_completion_message(False, error_msg)
            else:
                print(f"ERROR: {error_msg}")
            raise
    
    def _save_parsed_data(self, items_df, trip_number: Optional[str], tracking_number: Optional[str]):
        """
        Save parsed data to data/input/parsedInput.tsv.
        
        Args:
            items_df: DataFrame with parsed items
            trip_number: Trip identifier
            tracking_number: Tracking number
        """
        # Ensure data/input directory exists
        project_root = Path(__file__).parent.parent
        output_dir = project_root / "data" / "input"
        output_dir.mkdir(parents=True, exist_ok=True)
        
        output_file = output_dir / "parsedInput.tsv"
        
        # Add trip and tracking number as columns (fill all rows with same value)
        output_df = items_df.copy()
        output_df.insert(0, 'trip', trip_number or '')
        output_df.insert(1, 'tracking_number', tracking_number or '')
        
        # Save to TSV
        output_df.to_csv(output_file, sep='\t', index=False, encoding='utf-8')
        
        logger = get_logger(__name__)
        logger.info(f"Saved parsed data to {output_file}")
    
    def _search_production_numbers(self, items_df):
        """
        Search for production numbers using Enlabel automation.
        
        Args:
            items_df: DataFrame with item/lot combinations
        
        Returns:
            DataFrame with added production_number column
        """
        logger = get_logger(__name__)
        
        try:
            # Initialize Enlabel automation
            with EnlabelAutomation(self.config) as automation:
                # Login
                if self.gui:
                    self.gui.update_status("Logging in to Enlabel...")
                automation.login()
                
                # Search for production numbers
                if self.gui:
                    self.gui.update_status("Searching for production numbers...")
                result_df = automation.search_production_numbers(items_df)
                
                return result_df
        except Exception as e:
            logger.error(f"Error searching for production numbers: {e}", exc_info=True)
            if self.gui:
                self.gui.update_status(f"ERROR: Production number search failed: {str(e)}")
            # Return original dataframe with empty production_number column
            result_df = items_df.copy()
            result_df['production_number'] = None
            return result_df
    
    def _save_production_numbers(self, production_numbers_df, trip_number: Optional[str], tracking_number: Optional[str]):
        """
        Save production numbers to verification CSV file.
        
        Args:
            production_numbers_df: DataFrame with production numbers
            trip_number: Trip identifier
            tracking_number: Tracking number
        """
        # Ensure verification directory exists
        project_root = Path(__file__).parent.parent
        verification_dir = project_root / "data" / "verification"
        verification_dir.mkdir(parents=True, exist_ok=True)
        
        output_file = verification_dir / "production_numbers.csv"
        
        # Add trip and tracking number as columns
        output_df = production_numbers_df.copy()
        output_df.insert(0, 'trip', trip_number or '')
        output_df.insert(1, 'tracking_number', tracking_number or '')
        
        # Save to CSV
        output_df.to_csv(output_file, index=False, encoding='utf-8')
        
        logger = get_logger(__name__)
        logger.info(f"Saved production numbers to {output_file}")
    
    def run_gui(self):
        """Run the GUI application."""
        if self.gui is None:
            self.gui = FIFRAGUI()
        
        # Set callback for Start button
        self.gui.set_status_callback(self.process_files)
        
        # Run GUI
        logger = get_logger(__name__)
        logger.info("Starting GUI application")
        self.gui.run()


def main():
    """Main entry point."""
    import argparse
    
    parser = argparse.ArgumentParser(description="FIFRA Label Automation")
    parser.add_argument(
        '--gui',
        action='store_true',
        default=True,
        help='Run with GUI (default)'
    )
    parser.add_argument(
        '--no-gui',
        dest='gui',
        action='store_false',
        help='Run without GUI (command-line mode)'
    )
    parser.add_argument(
        '--tsv',
        type=str,
        help='Path to TSV file (command-line mode only)'
    )
    parser.add_argument(
        '--invoice',
        type=str,
        help='Path to invoice PDF (command-line mode only)'
    )
    
    args = parser.parse_args()
    
    automation = FIFRAAutomation()
    
    if args.gui:
        automation.run_gui()
    else:
        # Command-line mode
        if not args.tsv or not args.invoice:
            print("Error: --tsv and --invoice are required in command-line mode")
            sys.exit(1)
        automation.process_files(args.tsv, args.invoice)


if __name__ == "__main__":
    main()
