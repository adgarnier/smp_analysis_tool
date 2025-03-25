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
progress_data = {'current': 0, 'total': plate_amt, 'status': 'in-progress'}
p = 0

def main():
    if isinstance(plate_amt, int):
        with open(f'progresses/progress{user}.json', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        execute_signstatement()
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

def signing_statements():
    global x, tmp_password
    while True:
        try:
            wait_for_window(smp)
            time.sleep(1)
            # getting to the Statement window
            if x == 0 or x > 0:
                setting_tab()
                pyautogui.keyDown('shift')
                pyautogui.press('tab')
                pyautogui.keyUp('shift')
                for n in range (2):
                    logging.info(f'waiting {n+1}/2')
                    time.sleep(1)
                pyautogui.press('enter')                
                wait_for_window('Statements', timeout=30)
            else:
                logging.info(f'Error in {x}')

            # signing the Statement
            time.sleep(1)
            pyautogui.keyDown('shift')
            pyautogui.press('tab', presses = 3)
            time.sleep(1)
            pyautogui.press('tab', presses = 1)
            pyautogui.keyUp('shift')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(1)
            if find_window_by_title('Revoke Signature') or find_window_by_title('SoftMax Pro Application Help'):
                logging.info('Statement already signed. Quitting.')
                messagebox.showerror("Abort Status", "Statement already signed. Execution aborted.")
                sys.exit()              
            pyautogui.keyDown('shift')
            pyautogui.press('tab', presses = 2)
            pyautogui.keyUp('shift')
            time.sleep(2)
            pyautogui.typewrite('NA')
            time.sleep(1)
            pyautogui.press('tab', presses = 2)
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.typewrite(tmp_password)
            for n in range (2):
                time.sleep(1)
                logging.info(f'Waiting {n+1}/2 seconds')
            pyautogui.press('enter')
            time.sleep(1)

            # check if tmp_password is correct
            if find_window_by_title('Sign Statement'):
                logging.info('Incorrect password, please retry')
                messagebox.showerror("Abort Status", "Incorrect password, please retry. Execution aborted.")
                sys.exit()
            else:
                pyautogui.press('enter')

            # changing status to "Approval pending"
            time.sleep(3)
            # move to "In Progress"
            pyautogui.keyDown('shift')
            pyautogui.press('tab')
            pyautogui.keyUp('shift')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(1)
            wait_for_window('Document Status')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(2)
            break

        except pyautogui.FailSafeException:
            logging.info('PyAutoGUI fail-safe triggered from mouse moving to a corner of the screen. Stopping script.')
            sys.exit()
            
        except Exception as e:
            logging.info(f'An error occurred: {e}')
            break

def execute_signstatement():
    global x, p, tmp_password
    x = 0
    # popup box for password
    tmp_password = simpledialog.askstring("Input", "Enter your password:", show='*')
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
            signing_statements()
            time.sleep(1)
            progress()
        else:
            logging.info(f"Error at {x}")
            break
        x = x + 1
    logging.info(f"\nDone signing {x} plates.\n")

if __name__ == '__main__':
    main()