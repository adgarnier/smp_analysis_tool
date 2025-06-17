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
progress_data = {'current': 0, 'total': plate_amt, 'status': 'in-progress'}
p = 0

def main():
    if isinstance(plate_amt, int):
        with open(f'progresses/progress{user}.json', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        execute_cycle()
    else:
        if type(plate_amt) is not int:
            print('Error: <plate_amount> must be an integer.') 
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
    print('start')
    for n in range (plate_amt):
        print(f'{n}/{plate_amt}')
        time.sleep(1)
    print('end')
    
def execute_cycle():
    global x, p
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
            # no action
            time.sleep(0.5)
            progress()
        else:
            print(f'Error at {x}')
            break
        x = x + 1
    print(f'\nDone cycling {x} plates.\n')

if __name__ == '__main__':
    main()