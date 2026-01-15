"""
Test script for label verification functionality.
Tests verifying item number, lot number, and EPA number in label text.

This script can work with:
1. Text extracted from preview window (via pywinauto)
2. Text from a file (for testing verification logic)
3. Manual text input

Usage:
    python test_label_verification.py [--text "label text"] [--file path/to/text.txt] [--item ITEM] [--lot LOT]
    
Example:
    python test_label_verification.py --text "Item: NP6MSTGQP1 Lot: UE4376 EPA: 12345" --item NP6MSTGQP1 --lot UE4376
"""

import sys
import re
from pathlib import Path
from datetime import datetime
import argparse

# Logging setup
LOG_DIR = Path(__file__).parent
LOG_FILE = LOG_DIR / "test_label_verification_log.txt"


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


def verify_label_has_item_number(text, item_number, log_file):
    """
    Verify that label text contains the item number.
    
    Args:
        text: Label text to search
        item_number: Item number to verify
        log_file: Log file handle
    
    Returns:
        True if item number found, False otherwise
    """
    log(log_file, "\n" + "="*60)
    log(log_file, f"Verifying item number: {item_number}")
    log(log_file, "="*60)
    
    if not text:
        log(log_file, "No text provided for verification", "ERROR")
        return False
    
    if not item_number:
        log(log_file, "No item number provided for verification", "ERROR")
        return False
    
    # Normalize text and item number for comparison
    text_lower = text.lower()
    item_lower = item_number.lower().strip()
    
    log(log_file, f"Searching for item number '{item_number}' in label text...")
    log(log_file, f"Label text length: {len(text)} characters")
    log(log_file, f"Label text preview (first 200 chars): {text[:200]}...")
    
    # Method 1: Exact match (case-insensitive)
    if item_lower in text_lower:
        log(log_file, f"✓ Item number found (exact match, case-insensitive)")
        
        # Find all occurrences
        occurrences = []
        start = 0
        while True:
            pos = text_lower.find(item_lower, start)
            if pos == -1:
                break
            # Get context around the match
            context_start = max(0, pos - 20)
            context_end = min(len(text), pos + len(item_number) + 20)
            context = text[context_start:context_end]
            occurrences.append((pos, context))
            start = pos + 1
        
        log(log_file, f"Found {len(occurrences)} occurrence(s):")
        for i, (pos, context) in enumerate(occurrences):
            log(log_file, f"  Occurrence {i+1} at position {pos}: ...{context}...")
        
        return True
    
    # Method 2: Try with common prefixes/suffixes removed
    item_clean = item_lower.replace("item:", "").replace("item", "").strip()
    if item_clean and item_clean in text_lower:
        log(log_file, f"✓ Item number found (with prefix/suffix removed)")
        return True
    
    # Method 3: Try partial match (if item number is long)
    if len(item_number) > 6:
        # Try first 6 characters
        item_partial = item_lower[:6]
        if item_partial in text_lower:
            log(log_file, f"⚠ Partial match found (first 6 chars): {item_partial}", "WARNING")
            log(log_file, "This may be a false positive, manual verification recommended")
            return False
    
    log(log_file, f"✗ Item number '{item_number}' not found in label text", "ERROR")
    return False


def verify_label_has_lot_number(text, lot_number, log_file):
    """
    Verify that label text contains the lot number.
    
    Args:
        text: Label text to search
        lot_number: Lot number to verify
        log_file: Log file handle
    
    Returns:
        True if lot number found, False otherwise
    """
    log(log_file, "\n" + "="*60)
    log(log_file, f"Verifying lot number: {lot_number}")
    log(log_file, "="*60)
    
    if not text:
        log(log_file, "No text provided for verification", "ERROR")
        return False
    
    if not lot_number:
        log(log_file, "No lot number provided for verification", "ERROR")
        return False
    
    # Normalize text and lot number for comparison
    text_lower = text.lower()
    lot_lower = lot_number.lower().strip()
    
    log(log_file, f"Searching for lot number '{lot_number}' in label text...")
    log(log_file, f"Label text length: {len(text)} characters")
    
    # Method 1: Exact match (case-insensitive)
    if lot_lower in text_lower:
        log(log_file, f"✓ Lot number found (exact match, case-insensitive)")
        
        # Find all occurrences
        occurrences = []
        start = 0
        while True:
            pos = text_lower.find(lot_lower, start)
            if pos == -1:
                break
            # Get context around the match
            context_start = max(0, pos - 20)
            context_end = min(len(text), pos + len(lot_number) + 20)
            context = text[context_start:context_end]
            occurrences.append((pos, context))
            start = pos + 1
        
        log(log_file, f"Found {len(occurrences)} occurrence(s):")
        for i, (pos, context) in enumerate(occurrences):
            log(log_file, f"  Occurrence {i+1} at position {pos}: ...{context}...")
        
        return True
    
    # Method 2: Try with common prefixes/suffixes removed
    lot_clean = lot_lower.replace("lot:", "").replace("lot", "").replace("lote:", "").strip()
    if lot_clean and lot_clean in text_lower:
        log(log_file, f"✓ Lot number found (with prefix/suffix removed)")
        return True
    
    # Method 3: Try with spaces/hyphens normalized
    lot_normalized = re.sub(r'[\s\-_]', '', lot_lower)
    text_normalized = re.sub(r'[\s\-_]', '', text_lower)
    if lot_normalized in text_normalized:
        log(log_file, f"✓ Lot number found (with spaces/hyphens normalized)")
        return True
    
    log(log_file, f"✗ Lot number '{lot_number}' not found in label text", "ERROR")
    return False


def verify_label_has_epa_number(text, log_file):
    """
    Verify that label text contains an EPA number.
    We don't validate the format, just check for presence of "EPA" text.
    
    Args:
        text: Label text to search
        log_file: Log file handle
    
    Returns:
        True if EPA number found, False otherwise
    """
    log(log_file, "\n" + "="*60)
    log(log_file, "Verifying EPA number presence")
    log(log_file, "="*60)
    
    if not text:
        log(log_file, "No text provided for verification", "ERROR")
        return False
    
    text_lower = text.lower()
    
    log(log_file, "Searching for 'EPA' text in label...")
    log(log_file, f"Label text length: {len(text)} characters")
    
    # Method 1: Search for "EPA" text (case-insensitive)
    if "epa" in text_lower:
        log(log_file, "✓ EPA text found in label")
        
        # Find all occurrences and get context
        occurrences = []
        start = 0
        while True:
            pos = text_lower.find("epa", start)
            if pos == -1:
                break
            # Get context around the match
            context_start = max(0, pos - 30)
            context_end = min(len(text), pos + 30)
            context = text[context_start:context_end]
            occurrences.append((pos, context))
            start = pos + 1
        
        log(log_file, f"Found {len(occurrences)} occurrence(s) of 'EPA':")
        for i, (pos, context) in enumerate(occurrences):
            log(log_file, f"  Occurrence {i+1} at position {pos}: ...{context}...")
        
        # Try to extract EPA number if it follows "EPA"
        for pos, context in occurrences:
            # Look for patterns like "EPA 12345" or "EPA: 12345" or "EPA Reg. No. 12345"
            epa_patterns = [
                r'epa\s*:?\s*(\d+[-\d]*)',
                r'epa\s+reg\.?\s*no\.?\s*:?\s*(\d+[-\d]*)',
                r'epa\s+number\s*:?\s*(\d+[-\d]*)',
            ]
            for pattern in epa_patterns:
                match = re.search(pattern, context, re.IGNORECASE)
                if match:
                    epa_number = match.group(1)
                    log(log_file, f"  → Extracted EPA number: {epa_number}")
                    break
        
        return True
    
    # Method 2: Search for "EPA" with common variations
    epa_variations = ["epa reg", "epa registration", "epa number", "epa no"]
    for variation in epa_variations:
        if variation in text_lower:
            log(log_file, f"✓ EPA variation found: '{variation}'")
            return True
    
    log(log_file, "✗ EPA text not found in label", "ERROR")
    return False


def verify_label_complete(text, item_number, lot_number, log_file):
    """
    Perform complete label verification (item, lot, and EPA).
    
    Args:
        text: Label text to verify
        item_number: Expected item number
        lot_number: Expected lot number
        log_file: Log file handle
    
    Returns:
        dict with verification results
    """
    log(log_file, "\n" + "="*60)
    log(log_file, "Complete Label Verification")
    log(log_file, "="*60)
    log(log_file, f"Item Number: {item_number}")
    log(log_file, f"Lot Number: {lot_number}")
    log(log_file, "")
    
    results = {
        'item_number_found': False,
        'lot_number_found': False,
        'epa_number_found': False,
        'all_verified': False,
        'errors': []
    }
    
    # Verify item number
    results['item_number_found'] = verify_label_has_item_number(text, item_number, log_file)
    if not results['item_number_found']:
        results['errors'].append(f"Item number '{item_number}' not found")
    
    # Verify lot number
    results['lot_number_found'] = verify_label_has_lot_number(text, lot_number, log_file)
    if not results['lot_number_found']:
        results['errors'].append(f"Lot number '{lot_number}' not found")
    
    # Verify EPA number
    results['epa_number_found'] = verify_label_has_epa_number(text, log_file)
    if not results['epa_number_found']:
        results['errors'].append("EPA number not found")
    
    # Check if all verified
    results['all_verified'] = (
        results['item_number_found'] and
        results['lot_number_found'] and
        results['epa_number_found']
    )
    
    log(log_file, "")
    log(log_file, "="*60)
    log(log_file, "Verification Summary")
    log(log_file, "="*60)
    log(log_file, f"Item Number: {'✓ FOUND' if results['item_number_found'] else '✗ NOT FOUND'}")
    log(log_file, f"Lot Number: {'✓ FOUND' if results['lot_number_found'] else '✗ NOT FOUND'}")
    log(log_file, f"EPA Number: {'✓ FOUND' if results['epa_number_found'] else '✗ NOT FOUND'}")
    log(log_file, f"Overall: {'✓ VERIFIED' if results['all_verified'] else '✗ NOT VERIFIED'}")
    
    if results['errors']:
        log(log_file, "")
        log(log_file, "Errors:")
        for error in results['errors']:
            log(log_file, f"  - {error}", "ERROR")
    
    return results


def main():
    """Main test function."""
    parser = argparse.ArgumentParser(description="Test label verification functionality")
    parser.add_argument('--text', type=str, help='Label text to verify')
    parser.add_argument('--file', type=str, help='Path to file containing label text')
    parser.add_argument('--item', type=str, help='Item number to verify')
    parser.add_argument('--lot', type=str, help='Lot number to verify')
    
    args = parser.parse_args()
    
    log_file = setup_logging()
    
    log(log_file, "="*60)
    log(log_file, "Label Verification Test Script")
    log(log_file, "="*60)
    log(log_file, f"Test started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(log_file, "")
    
    # Get label text
    label_text = None
    
    if args.text:
        label_text = args.text
        log(log_file, "Using label text from command line argument")
    elif args.file:
        file_path = Path(args.file)
        if file_path.exists():
            log(log_file, f"Reading label text from file: {file_path}")
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    label_text = f.read()
                log(log_file, f"Read {len(label_text)} characters from file")
            except Exception as e:
                log(log_file, f"Error reading file: {e}", "ERROR")
                log_file.close()
                return
        else:
            log(log_file, f"File not found: {file_path}", "ERROR")
            log_file.close()
            return
    else:
        log(log_file, "No label text provided. Enter label text manually:")
        label_text = input("Label text: ").strip()
        if not label_text:
            log(log_file, "No label text entered, aborting test", "ERROR")
            log_file.close()
            return
    
    # Get item and lot numbers
    item_number = args.item
    lot_number = args.lot
    
    if not item_number:
        item_number = input("Enter item number to verify (or press Enter to skip): ").strip()
    
    if not lot_number:
        lot_number = input("Enter lot number to verify (or press Enter to skip): ").strip()
    
    log(log_file, "")
    log(log_file, f"Label text length: {len(label_text)} characters")
    log(log_file, f"Label text preview (first 500 chars):")
    log(log_file, label_text[:500])
    if len(label_text) > 500:
        log(log_file, f"... (truncated, total length: {len(label_text)} chars)")
    log(log_file, "")
    
    # Run verification tests
    if item_number and lot_number:
        # Complete verification
        results = verify_label_complete(label_text, item_number, lot_number, log_file)
    else:
        # Individual verifications
        if item_number:
            verify_label_has_item_number(label_text, item_number, log_file)
        
        if lot_number:
            verify_label_has_lot_number(label_text, lot_number, log_file)
        
        verify_label_has_epa_number(label_text, log_file)
    
    log(log_file, "")
    log(log_file, "="*60)
    log(log_file, "Test completed")
    log(log_file, "="*60)
    
    log_file.close()
    print(f"\nTest log saved to: {LOG_FILE}")


if __name__ == "__main__":
    main()
