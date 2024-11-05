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
setting_num = config['setting_num']
progress_data = {'current': 0, 'total': plate_amt, 'status': 'in-progress'}
p = 0

def main():
    if isinstance(plate_amt, int) and isinstance(setting_num, int):
        with open(f'progresses/progress{user}.json', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        execute_print()
    else:
        if type(plate_amt) is not int and type(setting_num) is not int:
            print('Error: <plate_amount> and <setting_num> must be an integers.')
            return
        elif type(plate_amt) is not int:
            print('Error: <plate_amount> must be an integer.') 
            return            
        elif type(setting_num) is not int:
            print('Error: <setting_num> must be an integer.') 
            return
        else:
            print('Error. Exiting script.')
            return
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
    for n in range (plate_amt):
        print(f'{n+1}/{plate_amt}')
        time.sleep(1)
    print('end')

def printing_plates():
    global x
    while True:
        try:
            wait_for_window(smp)
            pyautogui.keyDown('ctrl')
            pyautogui.press('p')
            pyautogui.keyUp('ctrl')
            wait_for_window('Printing from Win32 application')
            time.sleep(2)
            if setting_num == 86:
                pyautogui.press('tab', presses=7)
            else:    
                pyautogui.press('tab', presses=5)
                time.sleep(0.5)
                pyautogui.press('enter')
                wait_for_window('Printing Preferences')
                time.sleep(0.5)
                pyautogui.keyDown('shift')
                pyautogui.press('tab', presses=7)
                pyautogui.keyUp('shift')
                pyautogui.press('right', presses=setting_num)
            time.sleep(0.5)
            pyautogui.press('enter')
            wait_for_window('Printing from Win32 application')
            time.sleep(0.5)
            pyautogui.press('tab')
            time.sleep(2)
            pyautogui.press('enter')
            time.sleep(3)
            # print('Waiting 13 seconds for printing')
            # for n in range (13):
            #     print(f'{n+1}/13 seconds...')
            #     time.sleep(1)
            wait_for_window(smp, timeout=30)
            print(f'Plate {x+1} sent to printer.')
            break
        except pyautogui.FailSafeException:
            print('PyAutoGUI fail-safe triggered from mouse moving to a corner of the screen. Stopping script.')
            sys.exit()
        except Exception as e:
            print(f'An error occurred: {e}')
            break

def execute_print():
    global x, p
    x = 0
    window_focus(smp)
    while x < plate_amt:
        p = x
        wait_for_window(smp, timeout=30)
        print(f"Selecting file {x+1}")
        if x <= plate_amt:
            pyautogui.keyDown("ctrl")
            pyautogui.press("tab", presses=x)
            print(f"Tab + {x+1}")
            pyautogui.keyUp("ctrl")
            time.sleep(0.5)
            # action
            printing_plates()
            time.sleep(1)
            progress()
        else:
            print(f"Error at {x}")
            break
        x = x + 1
    print(f"\nDone printing {x} plates.\n")

if __name__ == '__main__':
    main()