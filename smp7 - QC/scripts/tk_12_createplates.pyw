import time
import os
import sys
import subprocess
import pyautogui
import win32gui
import json
import logging
from tkinter import simpledialog
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
setting_num = config['setting_num']
progress_data = {'current': 0, 'total': plate_amt, 'status': 'in-progress'}
p = 0

def main():
    if isinstance(plate_amt, int) and isinstance(setting_num, int):
        with open(f'progresses/progress{user}.json', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        execute_create_plates()
    else:
        if type(plate_amt) is not int and type(setting_num) is not int:
            logging.info('Error: <plate_amount> and <setting_num> must be an integers.')
            return
        elif type(plate_amt) is not int:
            logging.info('Error: <plate_amount> must be an integer.') 
            return            
        elif type(setting_num) is not int:
            logging.info('Error: <setting_num> must be an integer.') 
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
    global plate_amt
    logging.info('start')
    for n in range (plate_amt):
        logging.info(f'{n+1}/{plate_amt}')
        time.sleep(1)
    logging.info('end')

def create_plates():
    global x, full_plate_name
    plate_name = full_plate_name[:-4]
    plate_number = full_plate_name[-4:]
    while True:
        try:
            wait_for_window(smp)
            time.sleep(1)
            # write code here to change the plate number
            setting_tab()
            time.sleep(1)
            pyautogui.press('tab', presses=2)
            time.sleep(0.5)
            pyautogui.press('down', presses=4)
            tab
            right 8
            # saving files
            pyautogui.press(['alt', 'f', 'a'])
            wait_for_window('Open Dialog')
            time.sleep(0.5)
            pyautogui.press('tab')
            time.sleep(1)
            ### for clinical, plate is under the name PL-20250305-XXXX
            pyautogui.typewrite(plate_name)
            time.sleep(0.5)
            pyautogui.typewrite(plate_number)
            time.sleep(1)
            pyautogui.press('tab', presses=2)
            time.sleep(1)
            pyautogui.press('enter')
            break

        except pyautogui.FailSafeException:
            logging.info('PyAutoGUI fail-safe triggered from mouse moving to a corner of the screen. Stopping script.')
            sys.exit()
            
        except Exception as e:
            logging.info(f'An error occurred: {e}')
            break

def execute_create_plates():
    global x, p, global full_plate_name
    x = 0
    # popup box for plate name
    full_plate_name = simpledialog.askstring("Input", "Enter the name of your first plate:", )
    window_focus(smp)
    while x < plate_amt:
        p = x
        wait_for_window(smp, timeout=30)
        logging.info(f"Selecting file {x+1}")
        if x <= plate_amt:
            pyautogui.keyDown("ctrl")
            pyautogui.press("tab", presses=x)
            logging.info(f"Tab + {x+1}")
            pyautogui.keyUp("ctrl")
            time.sleep(0.5)
            # action
            create_plates()
            time.sleep(1)
            progress()
        else:
            logging.info(f"Error at {x}")
            break
        x = x + 1
    logging.info(f"\nDone signing {x} plates.\n")

if __name__ == '__main__':
    main()