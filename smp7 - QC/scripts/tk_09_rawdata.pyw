import time
import os
import sys
import subprocess
import pyautogui
import win32gui
import json
import logging
from tkinter import messagebox
from tk_00_base import *

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
plate_amt = config['plate_amt']
progress_data = {'current': 0, 'total': plate_amt, 'status': 'in-progress'}
p = 0

def main():
    if isinstance(plate_amt, int):
        with open(f'progresses/progress{user}.json', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        logging.info(f'Progress updated: {progress_data}')
        execute_rawdata()
    else:
        if type(plate_amt) is not int:
            logging.info('Error: <plate_amount> must be an integer.') 
            return            
        else:
            logging.info('Error. Exiting script.')
            return
    progress_data['status'] = 'done'
    with open(f'progresses/progress{user}.json', 'w') as progress_file:
        json.dump(progress_data, progress_file)
    window_focus(tool)
    logging.info('Commands complete.')

def progress():
    global p
    progress_data['current'] = p + 1
    try:
        with open(f'progresses/progress{user}.json', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        logging.info(f'Progress updated: {progress_data}')
    except Exception as e:
            logging.info(f'Error writing to progress.json: {e}')

def execute_test():
    global p
    p = 1
    logging.info(f'start {plate_amt}')
    for n in range (plate_amt):
        progress()
        p = p + 1
        logging.info(f'{n+1}/{plate_amt}')
        time.sleep(1)
    logging.info('end')

def rawdata():
    global x, p
    wait_for_window(smp)
    time.sleep(1)
    logging.info(f'Opening file {plate_amt-x}')
    # open file explorer
    pyautogui.hotkey('ctrlleft', 'o')
    time.sleep(0.5)
    wait_for_window('Open Dialog')
    time.sleep(0.5)

    # set open type
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('up')
    time.sleep(1)

    logging.info('\nLocate and open raw data file.\n')
    wait_for_window('SoftMax Pro', timeout=60)
    time.sleep(3)
    pyautogui.press('enter')

    # copy raw data
    logging.info('copying...')
    time.sleep(1)
    pyautogui.hotkey('ctrlleft', 'c')
    time.sleep(2)
    pyautogui.hotkey('ctrlleft', 'w')
    time.sleep(2)

    # paste raw data in template
    logging.info('pasting...')
    pyautogui.press('alt')
    time.sleep(1)
    pyautogui.press('v')
    time.sleep(1)
    pyautogui.press('o')
    for n in range(3):
        logging.info(f'Waiting {n+1}/3 seconds')
        time.sleep(1)
    if x == 0:
        for n in range(6):
            pyautogui.press('left')
            time.sleep(0.5)
        pyautogui.press('tab')
        time.sleep(2)
        pyautogui.press('left')
        time.sleep(2)
        pyautogui.press('right')
        time.sleep(2)
        pyautogui.press('left')
        time.sleep(2)
    if x > 0:
        for n in range(6):
            pyautogui.press('left')
            time.sleep(0.5)
        time.sleep(2)
        pyautogui.press('left')
        time.sleep(2)
        pyautogui.press('right')
        time.sleep(2)
        pyautogui.press('left')
    pyautogui.press('tab', presses=2)
    time.sleep(2)
    pyautogui.press('down', presses=3)
    time.sleep(2)
    pyautogui.press('right')
    time.sleep(2)
    pyautogui.press('down', presses=5)
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.press('alt')
    time.sleep(1)
    pyautogui.press('v')
    time.sleep(1)
    pyautogui.press('x')
    time.sleep(1)
    for n in range(3):
        logging.info(f'Waiting {n+1}/3 seconds')
        time.sleep(1)


    #     logging.info(f'Calling progress() with x = {x}, p = {p}')
    #     progress()
    # x = x + 1
    # logging.info(f'x incremented to {x}')

    logging.info(f'\nDone rawdata {x+1} files.\n')

def execute_rawdata():
    global x, p
    # each plate export to excel
    wait_for_window(smp)
    x = 0
    window_focus(smp)
    while x < plate_amt:
        p = x
        wait_for_window(smp)
        logging.info(f'Selecting file {x+1}')
        if x <= plate_amt:
            pyautogui.keyDown('ctrl')
            pyautogui.press('tab', presses=x)
            logging.info(f'tab + {x+1}')
            pyautogui.keyUp('ctrl')
            # action
            rawdata()
            time.sleep(3)
            progress()
        else:
            logging.info(f'Error at {x}')
            break
        x = x + 1
    logging.info(f'\nDone exporting {x+1} plates to excel.\n')

if __name__ == '__main__':
    main()