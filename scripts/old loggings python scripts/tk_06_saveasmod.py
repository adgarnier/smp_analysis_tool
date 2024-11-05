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
                logging.info('starting!')
            execute_save_as_mod()
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

def save_as_mod():
    global x
    while True:
        try:
            wait_for_window(smp)
            pyautogui.press(['alt', 'f', 'a'])
            wait_for_window('Save As')
            time.sleep(0.5)
            pyautogui.press('right')
            time.sleep(0.5)
            pyautogui.typewrite('_mod')
            time.sleep(0.5)
            pyautogui.press('tab', presses=3)
            time.sleep(0.1)
            pyautogui.press('enter')
            if find_window_title('SoftMax Pro'):
                logging.info(f'Cannot save plate {x+1}. Continuing with script...')
                pyautogui.press('enter')
                break                
            elif wait_for_window(smp):
                logging.info(f'Plate {x+1} saved as <plate name>_mod.')
                break
            else:
                logging.info(f'Error occured at {x+1}')
                sys.exit()
        except pyautogui.FailSafeException:
            logging.info('PyAutoGUI fail-safe triggered from mouse moving to a corner of the screen. Stopping script.')
            sys.exit()
        except Exception as e:
            logging.info(f'An error occurred: {e}')
            break

def execute_save_as_mod():
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
            # action
            save_as_mod()
            time.sleep(1)
            progress()
        else:
            logging.info(f'Error at {x}')
            break
        x = x + 1
    logging.info(f'\nDone saving {x} plates as <plate name>_mod.\n')

if __name__ == '__main__':
    main()