"""Main entry point for Excel AI Copilot desktop application."""

import sys
import os

# Ensure package is on python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_copilot.gui.app import main

if __name__ == "__main__":
    main()
