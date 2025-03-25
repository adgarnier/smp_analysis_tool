import time
import os
import sys
import subprocess
import pyautogui
import win32gui
import json
from tk_00_base import *
import logging

user = os.getlogin()
with open(f'configs/config{user}.json', 'r') as file:
    config = json.load(file)

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

# variables
# username = ''
# password = ''
plate_amt = 9
progress_data = {'current': 0, 'total': plate_amt, 'status': 'in-progress'}
x = 0

def main():
    try:
        global plate_amt
        global x
        execute_start()
        progress_data['status'] = 'done'
        with open(f'progresses/progress{user}.json', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        window_focus(tool)
        logging.info('Commands complete.')
    except Exception as e:
        logging.error(f'An unexpected error occurred: {e}', exc_info=True)

def progress():
    global x
    progress_data['current'] = x + 1
    with open(f'progresses/progress{user}.json', 'w') as progress_file:
        json.dump(progress_data, progress_file)
    x = x + 1

def execute_test():
    global x
    logging.info('start')
    for n in range (plate_amt):
        progress()
        logging.info(f'{n+1}/{plate_amt}')
        time.sleep(1)
        x = x + 1
    logging.info('end')

def login():
    global x
    logging.info(f'Starting {smp}...')
    try:
        subprocess.Popen(smp_location, creationflags=subprocess.CREATE_NO_WINDOW)
        progress()
        logging.info(f"{smp_location} is now open.")
    except FileNotFoundError:
        logging.info(f"Error: {smp_location} not found. Please verify the installation path.")
    logging.info('Waiting 10 seconds')
    for n in range (10): 
        logging.info(f'{n+1}/10 seconds...')
        time.sleep(1)
        if n % 2 == 0:
            progress()
    wait_for_window('SoftMax Pro GxP', timeout=60)
    window_focus('SoftMax Pro Gxp')
    progress()
    logging.info('\nSign-in using username and password.\n')
    # pyautogui.typewrite(username)
    # pyautogui.press('tab')
    # pyautogui.typewrite(password)
    # pyautogui.press('enter')
    time.sleep(3)
    progress()
    wait_for_window(smp, timeout=120)
    time.sleep(2)
    if find_window_by_title('Plate Setup Helper'):
        logging.info('Found plate setup helper')
        # wait_for_window('Plate Setup Helper', timeout=120)
        pyautogui.press('enter')
    else:
        logging.info('Plate setup helper not found')
    wait_for_window(smp)
    # time.sleep(2)
    pyautogui.getWindowsWithTitle(smp)[0].maximize()
    time.sleep(0.1)
    wait_for_window(smp)
    pyautogui.hotkey('ctrlleft', 'w')
    progress()
    logging.info(f'{smp} started.')

def execute_start():
    # opening the app and logging in
    login()
    logging.info('\nSuccessfully started.\n')

if __name__ == '__main__':
    main()