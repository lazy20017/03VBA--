# -*- coding: utf-8 -*-
"""
VBA Import Tool - Main Program Entry
Used for managing and importing/exporting VBA code from Word documents
"""
import sys
import os

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt5.QtWidgets import QApplication
from ui.main_window import MainWindow


def main():
    """Main function"""
    # Create application
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    # Set application info
    app.setApplicationName("VBA Import Tool")
    app.setApplicationVersion("1.0.0")

    # Create and show main window
    window = MainWindow()
    window.show()

    # Run application
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
