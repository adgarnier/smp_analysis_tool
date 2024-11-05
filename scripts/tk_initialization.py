import sys
import subprocess
from tkinter import messagebox

def main():
    check_and_install_pyautogui()
    check_and_install_pywin32()

def check_and_install_pyautogui():
    try:
        # Check if pyautogui is installed
        global pyautogui
        import pyautogui
        print('pyautogui is already installed.')
    except ImportError:
        # Install pyautogui using pip
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyautogui"])
            print("pyautogui has been successfully installed.")
        except subprocess.CalledProcessError:
            print("Error installing pyautogui. Please make sure you have pip installed.")

def check_and_install_pywin32():
    try:
        # Check if pywin32 is installed
        global win32gui
        import win32gui
        print('pywin32 is already installed.')
    except ImportError:
        # Install pywin32 using pip
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
            print("pywin32 has been successfully installed.")
        except subprocess.CalledProcessError:
            print("Error installing pywin32. Please make sure you have pip installed.")

if __name__ == '__main__':
    main()