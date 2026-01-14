"""
TSV file parser for Oracle ERP export files.
Extracts Trip, Tracking Number, Item Name, and Lot number data.
"""

import pandas as pd
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import re

from src.logger_setup import get_logger

logger = get_logger(__name__)


class TSVParser:
    """Parser for TSV files from Oracle ERP."""
    
    def __init__(self, config=None):
        """
        Initialize TSV parser.
        
        Args:
            config: Configuration object (optional, will use default if None)
        """
        if config is None:
            from src.config_loader import get_config
            config = get_config()
        
        self.config = config
        self.tsv_config = config.get_section('tsv')
        
    def parse_tsv(self, tsv_path: str) -> pd.DataFrame:
        """
        Parse TSV file and return DataFrame.
        
        Args:
            tsv_path: Path to TSV file
        
        Returns:
            DataFrame with parsed data
        """
        tsv_path = Path(tsv_path)
        if not tsv_path.exists():
            raise FileNotFoundError(f"TSV file not found: {tsv_path}")
        
        logger.info(f"Parsing TSV file: {tsv_path}")
        
        encoding = self.tsv_config.get('encoding', 'utf-8')
        delimiter = self.tsv_config.get('delimiter', '\t')
        
        try:
            df = pd.read_csv(
                tsv_path,
                sep=delimiter,
                encoding=encoding,
                dtype=str,  # Read all as strings to preserve formatting
                keep_default_na=False  # Don't convert empty strings to NaN
            )
            logger.info(f"Successfully parsed TSV file. Found {len(df)} rows.")
            return df
        except Exception as e:
            logger.error(f"Error parsing TSV file: {e}")
            raise
    
    def extract_key_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Extract key columns: Trip, Tracking Number, Item Name, Lot.
        
        Args:
            df: Input DataFrame
        
        Returns:
            DataFrame with only key columns
        """
        column_names = self.tsv_config['column_names']
        
        required_columns = [
            column_names['trip'],
            column_names['tracking_number'],
            column_names['item_name'],
            column_names['lot']
        ]
        
        # Check if all required columns exist
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing required columns in TSV file: {missing_columns}")
        
        # Extract key columns
        key_df = df[[
            column_names['trip'],
            column_names['tracking_number'],
            column_names['item_name'],
            column_names['lot']
        ]].copy()
        
        # Rename columns to standard names
        key_df.columns = ['trip', 'tracking_number', 'item_name', 'lot']
        
        logger.info(f"Extracted key columns. Found {len(key_df)} rows.")
        return key_df
    
    def filter_container_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Filter out container names (items starting with "CC-").
        These are not relevant for label processing.
        
        Args:
            df: DataFrame with item_name column
        
        Returns:
            DataFrame with container names excluded
        """
        initial_count = len(df)
        
        # Filter out rows where item_name starts with "CC-"
        filtered_df = df[~df['item_name'].str.strip().str.startswith('CC-', na=False)].copy()
        
        excluded_count = initial_count - len(filtered_df)
        if excluded_count > 0:
            logger.info(f"Excluded {excluded_count} container names (items starting with 'CC-').")
        
        return filtered_df
    
    def get_unique_items(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Get unique item/lot combinations, removing duplicates.
        Same item can have different lot numbers - keep all unique combinations.
        
        Args:
            df: DataFrame with item_name and lot columns
        
        Returns:
            DataFrame with unique item/lot combinations
        """
        # Keep unique combinations of item_name and lot
        unique_df = df[['item_name', 'lot']].drop_duplicates().copy()
        
        logger.info(f"Found {len(unique_df)} unique item/lot combinations.")
        return unique_df
    
    def get_trip_info(self, df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
        """
        Extract trip and tracking number information.
        Assumes all rows in the file belong to the same trip.
        
        Args:
            df: DataFrame with trip and tracking_number columns
        
        Returns:
            Tuple of (trip_number, tracking_number). Returns first non-empty values.
        """
        # Get first non-empty trip value
        trip_values = df['trip'].dropna()
        trip_values = trip_values[trip_values.str.strip() != '']
        
        # Get first non-empty tracking number value
        tracking_values = df['tracking_number'].dropna()
        tracking_values = tracking_values[tracking_values.str.strip() != '']
        
        trip_number = trip_values.iloc[0] if len(trip_values) > 0 else None
        tracking_number = tracking_values.iloc[0] if len(tracking_values) > 0 else None
        
        logger.info(f"Trip: {trip_number}, Tracking Number: {tracking_number}")
        return trip_number, tracking_number
    
    def validate_data(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, List[Dict]]:
        """
        Validate data and flag rows with missing/invalid entries.
        
        Args:
            df: DataFrame with item_name and lot columns
        
        Returns:
            Tuple of (valid_df, flagged_rows)
            - valid_df: DataFrame with valid rows
            - flagged_rows: List of dicts with flagged row info
        """
        flagged_rows = []
        valid_rows = []
        
        for idx, row in df.iterrows():
            item_name = str(row.get('item_name', '')).strip()
            lot = str(row.get('lot', '')).strip()
            
            issues = []
            if not item_name or item_name == '':
                issues.append("Missing item name")
            if not lot or lot == '':
                issues.append("Missing lot number")
            
            if issues:
                flagged_rows.append({
                    'index': idx,
                    'item_name': item_name,
                    'lot': lot,
                    'issues': issues
                })
            else:
                valid_rows.append(idx)
        
        valid_df = df.loc[valid_rows].copy()
        
        if flagged_rows:
            logger.warning(f"Found {len(flagged_rows)} rows with missing data that need manual confirmation.")
        
        return valid_df, flagged_rows
    
    def is_production_number(self, lot_number: str) -> bool:
        """
        Check if lot number is already a production number (9-digit number).
        
        Args:
            lot_number: Lot number string
        
        Returns:
            True if lot number matches 9-digit pattern
        """
        pattern = self.config.get('production_number.lot_is_production_number_pattern', r'^\d{9}$')
        return bool(re.match(pattern, str(lot_number).strip()))
    
    def parse_file(self, tsv_path: str) -> Dict:
        """
        Complete parsing workflow: parse file, extract columns, validate, get unique items.
        
        Args:
            tsv_path: Path to TSV file
        
        Returns:
            Dictionary with:
            - 'items': DataFrame with unique item/lot combinations
            - 'trip_number': Trip identifier
            - 'tracking_number': Tracking number
            - 'flagged_rows': List of rows that need manual confirmation
            - 'total_rows': Total number of rows in file
        """
        # Parse TSV file
        df = self.parse_tsv(tsv_path)
        total_rows = len(df)
        
        # Extract key columns
        key_df = self.extract_key_columns(df)
        
        # Filter out container names (items starting with "CC-")
        key_df = self.filter_container_names(key_df)
        
        # Get trip and tracking number
        trip_number, tracking_number = self.get_trip_info(key_df)
        
        # Validate data
        valid_df, flagged_rows = self.validate_data(key_df)
        
        # Get unique item/lot combinations
        unique_items = self.get_unique_items(valid_df)
        
        # Add production number check flag
        unique_items['is_production_number'] = unique_items['lot'].apply(self.is_production_number)
        
        result = {
            'items': unique_items,
            'trip_number': trip_number,
            'tracking_number': tracking_number,
            'flagged_rows': flagged_rows,
            'total_rows': total_rows
        }
        
        logger.info(f"Parsing complete. {len(unique_items)} unique items, {len(flagged_rows)} flagged rows.")
        return result
