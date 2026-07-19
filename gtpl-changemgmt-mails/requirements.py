import subprocess
import sys
import importlib

# List of required modules
required_modules = ['pandas','bs4','openpyxl']

# Function to check if Python is installed
def check_python():
    try:
        # Check Python version
        python_version = sys.version
        print(f"\nPython is installed: {python_version}\n")
    except Exception:
        print("\nPython is not installed. Please install Python and try again.\n")
        exit()

# Function to check and install required modules
def check_and_install_modules():
    for module in required_modules:
        try:
            importlib.import_module(module)
            print(f"\nModule '{module}' is already installed.\n")
        except ImportError:
            print(f"\nModule '{module}' is not installed. Installing...\n")
            try:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', module])
                print(f"\nModule '{module}' has been installed successfully.\n")
            except Exception as e:
                print(f"\nFailed to install module '{module}'. Reason: {e}\n")
                exit()

# Function to check for win32com (part of pywin32)
def check_win32com():
    try:
        import win32com.client
        print("\nPackage 'pywin32' is already installed.\n")
    except ImportError:
        print("\nPackage 'pywin32' is not installed. Installing pywin32...\n")
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pywin32'])
            print("\nPackage 'pywin32' has been installed successfully.\n")
        except Exception as e:
            print(f"\nFailed to install 'pywin32'. Reason: {e}\n")
            exit()

# Check if Python is installed
check_python()

# Check and install required modules
check_and_install_modules()

# Check for win32com module
check_win32com()
