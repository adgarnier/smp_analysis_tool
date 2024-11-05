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
# username = ''
# password = ''
plate_amt = 9
progress_data = {'current': 0, 'total': plate_amt, 'status': 'in-progress'}
x = 0

def main():
    global plate_amt
    global x
    execute_start()
    progress_data['status'] = 'done'
    with open(f'progresses/progress{user}.json', 'w') as progress_file:
        json.dump(progress_data, progress_file)
    window_focus(tool)
    print('Commands complete.')

def progress():
    global x
    progress_data['current'] = x + 1
    with open(f'progresses/progress{user}.json', 'w') as progress_file:
        json.dump(progress_data, progress_file)
    x = x + 1

def execute_test():
    global x
    print('start')
    for n in range (plate_amt):
        progress()
        print(f'{n+1}/{plate_amt}')
        time.sleep(1)
        x = x + 1
    print('end')

def login():
    global x
    print(f'Starting {smp}...')
    try:
        subprocess.Popen(smp_location)
        progress()
        print(f"{smp_location} is now open.")
    except FileNotFoundError:
        print(f"Error: {smp_location} not found. Please verify the installation path.")
    print('Waiting 10 seconds')
    for n in range (10): 
        print(f'{n+1}/10 seconds...')
        time.sleep(1)
        if n % 2 == 0:
            progress()
    wait_for_window('SoftMax Pro GxP', timeout=60)
    window_focus('SoftMax Pro Gxp')
    progress()
    print('\nSign-in using username and password.\n')
    # pyautogui.typewrite(username)
    # pyautogui.press('tab')
    # pyautogui.typewrite(password)
    # pyautogui.press('enter')
    time.sleep(3)
    progress()
    wait_for_window(smp, timeout=120)
    time.sleep(5)
    if find_window_by_title('Plate Setup Helper'):
        print('Found plate setup helper')
        # wait_for_window('Plate Setup Helper', timeout=120)
        pyautogui.press('enter')
    else:
        print('Plate setup helper not found')
    wait_for_window(smp)
    # time.sleep(2)
    pyautogui.getWindowsWithTitle(smp)[0].maximize()
    time.sleep(0.1)
    wait_for_window(smp)
    pyautogui.hotkey('ctrlleft', 'w')
    progress()
    print(f'{smp} started.')

def execute_start():
    # opening the app and logging in
    login()
    print('\nSuccessfully started.\n')

if __name__ == '__main__':
    main()