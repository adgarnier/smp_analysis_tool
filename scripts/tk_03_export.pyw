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
export_setting = None
export_setting2 = None
save_location = None
try:
    export_setting = config['export_setting']
except:
    logging.info(f'no export_setting saved {export_setting}.')
try:
    export_setting2 = config['export_setting2']
except:
    logging.info(f'no export_setting2 saved {export_setting2}.')
try:
    save_location = config[r'directory']
except:
    logging.info(f'no save_location saved {save_location}.')
lims_location = '\\\\morsqqlbv01p\\LIMS_DATA_IMPORT_PROD'
progress_data = {'current': 0, 'total': plate_amt, 'status': 'in-progress'}
p = 0

def main():
    try:
        with open(f'progresses/progress{user}.json', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        if 0 in export_setting or 7 in export_setting2:
            logging.info(f'Doing test {execute_export} excel')
            execute_export('excel')
        elif 1 in export_setting or 6 in export_setting2:
            logging.info(f'Doing test {execute_export} lims')
            execute_export('lims')
        elif 2 in export_setting or 8 in export_setting2:
            logging.info(f'Doing test {execute_export} pnps')
            execute_export('pnps')
        else:
            logging.info(f'Invalid export_setting: {export_setting} or {export_setting2}.')
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
    global plate_amt
    logging.info('start')
    logging.info(save_location)
    for n in range (plate_amt):
        logging.info(f'{n+1}/{plate_amt}')
        time.sleep(1)
    logging.info('end')

def export_selection(*options):
    # get to data spot
    wait_for_window(smp)
    pyautogui.press(['alt', 'f', 'e'])
    pyautogui.keyDown('shift')
    pyautogui.press('tab', presses=2)
    pyautogui.keyUp('shift')
    wait_for_window('Export')
    # select specfics
    for option in options:
        if option == 'excel':
            logging.info('Exporting excel file...')
            # export excel
            pyautogui.press(['down', 'right', 'space'])
            time.sleep(0.5)
            pyautogui.press('tab', presses=7)
            time.sleep(0.5)
            pyautogui.press('enter')
        elif option == 'lims':
            logging.info('Exporting to lims...')
            # export lims
            pyautogui.press(['end', 'right', 'space'])
            time.sleep(0.5)
            pyautogui.press('tab')
            time.sleep(0.5)
            pyautogui.press('enter')
        elif option == 'pnps':
            logging.info('Exporting pnps...')
            # export pnps
            # 3
            for n in range (4):
                pyautogui.press(['down', 'right'])
            pyautogui.press('space')
            # 6
            for n in range (3):
                pyautogui.press(['down', 'right'])
            pyautogui.press('space')           
            # 8
            for n in range (2):
                pyautogui.press(['down', 'right'])
            pyautogui.press('space')            
            # 9
            for n in range (1):
                pyautogui.press(['down', 'right'])
            pyautogui.press('space') 
            time.sleep(0.5)
            pyautogui.press('tab')
            time.sleep(0.5)
            pyautogui.press('enter')
            break
        else:
            logging.info(f'Invalid option: {option}')
            break

def export(*options):
    global x
    while True:
        try:
            # select export options
            for option in options:
                if option == 'excel':
                    logging.info(f'Selecting export options for excel')
                    export_selection('excel')
                elif option == 'lims':
                    logging.info(f'Selecting export options to lims')
                    export_selection('lims')
                elif option == 'pnps':
                    logging.info('Selecting export option pnps')
                    export_selection('pnps')
                else:
                    logging.info(f'Invalid option: {option}')
                    break 
            if wait_for_window('Save As'):
                if x == 0:
                    time.sleep(1)
                    pyautogui.press(['tab', 'right', 'right', 'enter'])
                    # edit address to Share
                    time.sleep(1.5)
                    pyautogui.keyDown('ctrl')
                    pyautogui.press('l')
                    pyautogui.keyUp('ctrl')
                    time.sleep(1)
                    for option in options:
                        if option == 'excel' or option == 'pnps':
                            logging.info(f'Saving to {save_location}')
                            pyautogui.typewrite(save_location)
                        elif option == 'lims':
                            logging.info(f'Saving to {lims_location}')
                            pyautogui.typewrite(lims_location)
                            time.sleep(1)
                        else:
                            logging.info(f'Invalid option: {option}')
                            break
                    time.sleep(1)
                    pyautogui.press('enter')
                    time.sleep(1.5)
                    # save first
                    pyautogui.press('tab', presses=9)
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(0.5)
                    if find_window_by_title('Confirm Save As'):
                        logging.info(f'Plate {x+1} already exported, overriding')
                        time.sleep(1)
                        pyautogui.press(['left', 'enter'])
                        break    
                    elif wait_for_window(smp):
                        logging.info(f'Done export for plate {x+1}')
                        break
                    else:
                        logging.info('Restarting export due to failure in activating SoftMax Pro window...')
                        
                if x > 0:
                    time.sleep(1.5)
                    pyautogui.press(['tab', 'right', 'right', 'enter'])
                    time.sleep(1)
                    # save
                    pyautogui.press('tab', presses=2)
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    time.sleep(0.5)
                    if find_window_by_title('Confirm Save As'):
                        logging.info(f'Plate {x+1} already exported, overriding')
                        time.sleep(1)
                        pyautogui.press(['left', 'enter'])
                        break    
                    elif wait_for_window(smp):
                        logging.info(f'Done export for plate {x+1}')
                        break
                    else:
                        logging.info('Restarting export due to failure in activating SoftMax Pro window...')
                else:
                    logging.info(f'Error in {x}')
            else:
                logging.info('Restarting export due to failure in activating Save As window...')
        except pyautogui.FailSafeException:
            logging.info('PyAutoGUI fail-safe triggered from mouse moving to a corner of the screen. Stopping script.')
            sys.exit()
        except Exception as e:
            logging.info(f'An error occurred: {e}')
            break

def execute_export(option):
    global x, p
    # each plate export to excel
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
            export(option)
            time.sleep(1)
            progress()
        else:
            logging.info(f'Error at {x}')
            break
        x = x + 1
    logging.info(f'\nDone exporting {x} plates to excel.\n')

if __name__ == '__main__':
    main()