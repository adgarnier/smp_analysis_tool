import time
import os
import sys
import subprocess
import pyautogui
import win32gui
import json
from tk_00_base import *

user = os.getlogin()
with open(f'configs/config{user}.json', 'r') as file:
    config = json.load(file)

# variables
plate_amt = config['plate_amt']
export_setting = None
export_setting2 = None
save_location = None
try:
    export_setting = config['export_setting']
except:
    print(f'no export_setting saved {export_setting}.')
try:
    export_setting2 = config['export_setting2']
except:
    print(f'no export_setting2 saved {export_setting2}.')
try:
    save_location = config[r'directory']
except:
    print(f'no save_location saved {save_location}.')
lims_location = '\\\\morsqqlbv01p\\LIMS_DATA_IMPORT_PROD'
progress_data = {'current': 0, 'total': plate_amt, 'status': 'in-progress'}
p = 0

def main():
    with open(f'progresses/progress{user}.json', 'w') as progress_file:
        json.dump(progress_data, progress_file)
    if 0 in export_setting or 7 in export_setting2:
        print(f'Doing test {execute_export} excel')
        execute_export('excel')
    elif 1 in export_setting or 6 in export_setting2:
        print(f'Doing test {execute_export} lims')
        execute_export('lims')
    elif 2 in export_setting or 8 in export_setting2:
        print(f'Doing test {execute_export} pnps')
        execute_export('pnps')
    else:
        print(f'Invalid export_setting: {export_setting} or {export_setting2}.')
    progress_data['status'] = 'done'
    with open(f'progresses/progress{user}.json', 'w') as progress_file:
        json.dump(progress_data, progress_file)        
    window_focus(tool)
    print('Commands complete.')

def progress():
    global p
    progress_data['current'] = p + 1
    try:
        with open(f'progresses/progress{user}.json', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        print(f'Progress updated: {progress_data}')
    except Exception as e:
            print(f'Error writing to progress.json: {e}')

def execute_test():
    global plate_amt
    print('start')
    print(save_location)
    for n in range (plate_amt):
        print(f'{n+1}/{plate_amt}')
        time.sleep(1)
    print('end')

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
            print('Exporting excel file...')
            # export excel
            pyautogui.press(['down', 'right', 'space'])
            time.sleep(0.5)
            pyautogui.press('tab', presses=7)
            time.sleep(0.5)
            pyautogui.press('enter')
        elif option == 'lims':
            print('Exporting to lims...')
            # export lims
            pyautogui.press(['end', 'right', 'space'])
            time.sleep(0.5)
            pyautogui.press('tab')
            time.sleep(0.5)
            pyautogui.press('enter')
        elif option == 'pnps':
            print('Exporting pnps...')
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
            print(f'Invalid option: {option}')
            break

def export(*options):
    global x
    while True:
        try:
            # select export options
            for option in options:
                if option == 'excel':
                    print(f'Selecting export options for excel')
                    export_selection('excel')
                elif option == 'lims':
                    print(f'Selecting export options to lims')
                    export_selection('lims')
                elif option == 'pnps':
                    print('Selecting export option pnps')
                    export_selection('pnps')
                else:
                    print(f'Invalid option: {option}')
                    break 
            if wait_for_window('Save As'):
                if x == 0:
                    time.sleep(0.5)
                    pyautogui.press(['tab', 'right', 'right', 'enter'])
                    # edit address to Share
                    pyautogui.keyDown('ctrl')
                    pyautogui.press('l')
                    pyautogui.keyUp('ctrl')
                    time.sleep(1)
                    for option in options:
                        if option == 'excel' or option == 'pnps':
                            print(f'Saving to {save_location}')
                            pyautogui.typewrite(save_location)
                        elif option == 'lims':
                            print(f'Saving to {lims_location}')
                            pyautogui.typewrite(lims_location)
                        else:
                            print(f'Invalid option: {option}')
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
                        print(f'Plate {x+1} already exported, overriding')
                        time.sleep(1)
                        pyautogui.press(['left', 'enter'])
                        break    
                    elif wait_for_window(smp):
                        print(f'Done export for plate {x+1}')
                        break
                    else:
                        print('Restarting export due to failure in activating SoftMax Pro window...')
                        
                if x > 0:
                    time.sleep(0.5)
                    pyautogui.press(['tab', 'right', 'right', 'enter'])
                    # save
                    pyautogui.press('tab', presses=2)
                    pyautogui.press('enter')
                    if find_window_by_title('Confirm Save As'):
                        print(f'Plate {x+1} already exported, overriding')
                        time.sleep(1)
                        pyautogui.press(['left', 'enter'])
                        break    
                    elif wait_for_window(smp):
                        print(f'Done export for plate {x+1}')
                        break
                    else:
                        print('Restarting export due to failure in activating SoftMax Pro window...')
                else:
                    print(f'Error in {x}')
            else:
                print('Restarting export due to failure in activating Save As window...')
        except pyautogui.FailSafeException:
            print('PyAutoGUI fail-safe triggered from mouse moving to a corner of the screen. Stopping script.')
            sys.exit()
        except Exception as e:
            print(f'An error occurred: {e}')
            break

def execute_export(option):
    global x, p
    # each plate export to excel
    x = 0
    window_focus(smp)
    while x < plate_amt:
        p = x
        wait_for_window(smp)
        print(f'Selecting file {x+1}')
        if x <= plate_amt:
            pyautogui.keyDown('ctrl')
            pyautogui.press('tab', presses=x)
            print(f'tab + {x+1}')
            pyautogui.keyUp('ctrl')
            # action
            export(option)
            time.sleep(1)
            progress()
        else:
            print(f'Error at {x}')
            break
        x = x + 1
    print(f'\nDone exporting {x} plates to excel.\n')

if __name__ == '__main__':
    main()