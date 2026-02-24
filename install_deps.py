# Install required packages
import subprocess
import sys

packages = ['pywin32', 'PyQt5', 'python-docx']

for package in packages:
    print(f"Installing {package}...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])

print("All packages installed successfully!")
