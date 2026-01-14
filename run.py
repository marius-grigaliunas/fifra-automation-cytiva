#!/usr/bin/env python
"""
Entry point for FIFRA Automation.
Run this script from the project root directory.
"""

import sys
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# Now import and run main
if __name__ == "__main__":
    from src.main import main
    main()
