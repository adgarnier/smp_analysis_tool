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
plate_amt = config['plate_amt']
progress_data = {'current': 0, 'total': plate_amt, 'status': 'in-progress'}
p = 0

def main():
    try:
        if isinstance(plate_amt, int):
            with open(f'progresses/progress{user}.json', 'w') as progress_file:
                json.dump(progress_data, progress_file)
            execute_cycle()
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
    logging.info('start')
    for n in range (plate_amt):
        logging.info(f'{n}/{plate_amt}')
        time.sleep(1)
    logging.info('end')
    
def execute_cycle():
    global x, p
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
            # no action
            time.sleep(0.5)
            progress()
        else:
            logging.info(f'Error at {x}')
            break
        x = x + 1
    logging.info(f'\nDone cycling {x} plates.\n')

if __name__ == '__main__':
    main()