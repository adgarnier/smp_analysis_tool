import time
import os
import sys
import subprocess
import pyautogui
import win32gui
import json
from tkinter import messagebox
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
plate_amt = config['plate_amt']
file_amt = config['file_amt']
progress_data = {'current': 0, 'total': file_amt, 'status': 'in-progress'}
p = 0

def main():
    try:
        if isinstance(plate_amt, int) and isinstance(file_amt, int):
            with open(f'progresses/progress{user}.json', 'w') as progress_file:
                json.dump(progress_data, progress_file)
            logging.info(f'Progress updated: {progress_data}')
            execute_open()
        else:
            if type(plate_amt) is not int and type(file_amt) is not int:
                logging.info('Error: <plate_amount> and <file_amount> must be an integers.')
                return
            elif type(plate_amt) is not int:
                logging.info('Error: <plate_amount> must be an integer.') 
                return            
            elif type(file_amt) is not int:
                logging.info('Error: <file_amount> must be an integer.') 
                return
            else:
                logging.info('Error. Exiting script.')
                return
        progress_data['status'] = 'done'
        with open(f'progresses/progress{user}.json', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        window_focus(tool)
        logging.info('Commands complete.')
    except Exception as e:
        logging.error(f'An unexpected error occurred: {e}', exc_info=True)

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
    logging.info(f'start {plate_amt} {file_amt}')
    for n in range (plate_amt):
        progress()
        p = p + 1
        logging.info(f'{n+1}/{plate_amt}')
        time.sleep(1)
    logging.info('end')

def execute_open():
    global x, p
    window_focus(smp)
    # open file explorer
    pyautogui.keyDown('ctrl')
    pyautogui.press('o')
    pyautogui.keyUp('ctrl')
    wait_for_window('Open')
    
    logging.info('\nLocate and open final plate of run.\n')
    # ensure smp is selected
    wait_for_window(smp, timeout=60)
    time.sleep(1)
    pyautogui.press('enter')
    progress()
    
    x = 1
    while x < file_amt:
        p = x
        logging.info(f'Opening file {file_amt-x}')
        # open file explorer
        pyautogui.keyDown('ctrl')
        pyautogui.press('o')
        pyautogui.keyUp('ctrl')
        time.sleep(0.5)
        wait_for_window('Open')
        time.sleep(0.5)
    
        # navigate to data
        pyautogui.press('tab', presses=9)
        pyautogui.press('end')
    
        # select and open file
        if x <= file_amt:
            pyautogui.press('up', presses=x)
            logging.info(f'up + {x+1}')
            time.sleep(0.1)
            pyautogui.press('enter')
            wait_for_window('SoftMax Pro')
            time.sleep(3)
            pyautogui.press('enter')
            time.sleep(1)
            logging.info(f'Calling progress() with x = {x}, p = {p}')
            progress()
        x = x + 1
        logging.info(f'x incremented to {x}')

    logging.info(f'\nDone opening {x} files.\n')
   
    if file_amt == plate_amt:
        return
    else:
        messagebox.showinfo('Info', 'Please close the unnecessary files.')
        # required_keyword = "ok"
        # logging.info(f'\nClose the unnecessary files and then type "ok" in the Terminal to continue with the script.\n')
        # user_input = input(f'Please enter "ok" to continue: ')
        # if user_input.lower() == required_keyword:
        #     logging.info('Keyword accepted. Continuing with the script...')
        #     window_focus(smp)
        # else:
        #     logging.info("Incorrect keyword. Script terminated.")
        #     sys.exit()

if __name__ == '__main__':
    main()