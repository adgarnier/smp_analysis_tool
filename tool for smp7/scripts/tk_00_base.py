## FULL SCRIPT
import tkinter as tk
from tkinter import messagebox
import time
import sys
import pyautogui
import win32gui
import logging
import os

smp = 'SoftMax Pro 7.1.2.1 GxP'
smp_location = r'C:\Program Files (x86)\Molecular Devices\SoftMax Pro 7.1.2\SoftMaxProApp.exe'
lims_location = r'\\morsqqlbv01p\LIMS_DATA_IMPORT_PROD'
tool = 'SMP Analysis Tool v7.1.2'

user = os.getlogin()
# Ensure the logs directory exists
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

def main():
    if len(sys.argv) < 4:
        logging.info('\nUsage: python smp_script.py <plate_amount> <file_amount> <command1> <command2> ...')
        logging.info(' plate_amount is the number of plates for the experiment.')
        logging.info(' file_amount is the number of files saved for the experiment (includes mods).')
        logging.info(' Commands: start, open, print, excel, lims, save_as_mod, full, masking, cycle\n')
        return
    else:
        try:
            global plate_amt
            plate_amt = int(sys.argv[1])
        except ValueError:
            logging.info('Error: <plate_amount> must be an integer')
            return
        try:
            global file_amt
            file_amt = int(sys.argv[2])
        except ValueError:
            logging.info('Error: <plate_amount> must be an integer')
            return

        commands = sys.argv[3:]
        valid_commands = {'start': execute_start, 'open': execute_open, 'print': execute_print, 'excel': execute_excel, \
                          'lims': execute_lims, 'full': execute_full, 'save_as_mod': execute_save_as_mod, \
                          'masking': execute_masking, 'cycle': execute_cycle}

        for command in commands:
            if command in valid_commands:
                valid_commands[command]()
            else:
                logging.info(f'Invalid command: {command}. Please use one of more of the following: start, open, print, excel, lims, full')
                return
    window_focus('Windows PowerShell')
    logging.info('Commands complete.')

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
          
def setting_tab():
    logging.info('Setting tab...')
    pyautogui.press(['alt', 'v', 'n'])
    time.sleep(0.5)
    pyautogui.press(['alt', 'v', 'o'])
    time.sleep(0.5)
    for n in range (8):
        pyautogui.press('left')
        logging.info(f'{n+1} left')
        time.sleep(0.3)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('left')
    time.sleep(1)
    pyautogui.press('right')
    time.sleep(1)
    pyautogui.press('left')
    time.sleep(1)
    pyautogui.press(['alt', 'v', 'n'])
    time.sleep(0.5)
    pyautogui.press(['alt', 'v', 'x'])
    time.sleep(2)
    logging.info('tab set to whole area')

if __name__ == '__main__':
    main()