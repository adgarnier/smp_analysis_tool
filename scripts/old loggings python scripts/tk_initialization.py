import sys
import subprocess
from tkinter import messagebox
import os
import logging

# Ensure the logs directory exists
user = os.getlogin()
log_directory = os.path.join('progresses', 'debugging')
if not os.path.exists(log_directory):
    os.makedirs(log_directory)

# Logging setup
log_file_path = os.path.join(log_directory, f'{user}_log.log')
logging.basicConfig(filename=log_file_path, level=logging.DEBUG, 
                    format='%(asctime)s %(levelname)s:%(message)s')
script_name = os.path.basename(__file__)
logging.info("\n" + "="*50 + f"\nRunning {script_name}\n" + "="*50)

def main():
    try:
        check_and_install_pyautogui()
        check_and_install_pywin32()
    except Exception as e:
        logging.error(f'An unexpected error occurred: {e}', exc_info=True)

def check_and_install_pyautogui():
    try:
        # Check if pyautogui is installed
        global pyautogui
        import pyautogui
        logging.info('pyautogui is already installed.')
    except ImportError:
        # Install pyautogui using pip
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyautogui"])
            logging.info("pyautogui has been successfully installed.")
        except subprocess.CalledProcessError:
            logging.info("Error installing pyautogui. Please make sure you have pip installed.")

def check_and_install_pywin32():
    try:
        # Check if pywin32 is installed
        global win32gui
        import win32gui
        logging.info('pywin32 is already installed.')
    except ImportError:
        # Install pywin32 using pip
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
            logging.info("pywin32 has been successfully installed.")
        except subprocess.CalledProcessError:
            logging.info("Error installing pywin32. Please make sure you have pip installed.")

if __name__ == '__main__':
    main()