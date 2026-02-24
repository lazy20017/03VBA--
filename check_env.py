# -*- coding: utf-8 -*-
"""Test script to check environment"""
import subprocess
import sys

print("Python version:", sys.version)

# Check installed packages
packages_to_check = ['PyQt5', 'win32com', 'docx']
for pkg in packages_to_check:
    try:
        __import__(pkg)
        print(f"{pkg}: OK")
    except ImportError:
        print(f"{pkg}: NOT FOUND")
