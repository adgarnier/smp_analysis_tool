## FULL SCRIPT
import tkinter as tk
from tkinter import messagebox
import time
import sys
import pyautogui
import win32gui
import os
import logging

tool = 'SMP Analysis Tool v3.1'
smp = 'SoftMax Pro 6.5.1 SP2 GxP'
smp_location = r'C:\Program Files (x86)\Molecular Devices\SoftMax Pro 6.5.1 GxP\SoftMaxProApp.exe'
lims_location = r'\\morsqqlbv01p\LIMS_DATA_IMPORT_PROD'

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

# main scripts

def wait_for_window(window_title, timeout=20):
    start_time = time.time()
    while True:
        if time.time() - start_time > timeout:
            logging.info(f'Timed out after {timeout} seconds, while waiting for {window_title} window...')
            messagebox.showerror("Abort Status", "Timed out. Execution aborted.")
            sys.exit()
            return False
        try:
            window = pyautogui.getWindowsWithTitle(window_title)[0]
            if window.isActive:
                logging.info(f'{window_title} window is now active. Continuing with the script...')
                return True
        except IndexError:
            try:
                kill = pyautogui.getWindowsWithTitle('Windows Powershell')[0]
                if kill.isActive:
                    logging.info(f'Kill activated. Exiting script.')
                    sys.exit()
                else:
                    logging.info(f'Waiting for {window_title} window...')
                    time.sleep(1)
            except IndexError:
                logging.info(f'Waiting for {window_title}...')
                time.sleep(1)

def find_window_by_title(title):
    hwnd = win32gui.FindWindow(None, title)
    return hwnd

def window_focus(window_title):
    window_handle = find_window_by_title(window_title)
    if window_handle:
        # Set the window to the foreground
        win32gui.SetForegroundWindow(window_handle)
        logging.info(f'Window \'{window_title}\' is now in focus.')
    else:
        logging.info(f'Window \'{window_title}\' not found.')
          
if __name__ == '__main__':
    main()