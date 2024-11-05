import time
import os
import sys
import subprocess
import pyautogui
import win32gui
import json
from tkinter import messagebox
from tk_00_base import *

user = os.getlogin()
with open(f'configs/config{user}.json', 'r') as file:
    config = json.load(file)

# variables
plate_amt = config['plate_amt']
file_amt = config['file_amt']
progress_data = {'current': 0, 'total': file_amt, 'status': 'in-progress'}
p = 0

def main():
    if isinstance(plate_amt, int) and isinstance(file_amt, int):
        with open(f'progresses/progress{user}.json', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        print(f'Progress updated: {progress_data}')
        execute_open()
    else:
        if type(plate_amt) is not int and type(file_amt) is not int:
            print('Error: <plate_amount> and <file_amount> must be an integers.')
            return
        elif type(plate_amt) is not int:
            print('Error: <plate_amount> must be an integer.') 
            return            
        elif type(file_amt) is not int:
            print('Error: <file_amount> must be an integer.') 
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
    global p
    p = 1
    print(f'start {plate_amt} {file_amt}')
    for n in range (plate_amt):
        progress()
        p = p + 1
        print(f'{n+1}/{plate_amt}')
        time.sleep(1)
    print('end')

def execute_open():
    global x, p
    window_focus(smp)
    # open file explorer
    pyautogui.keyDown('ctrl')
    pyautogui.press('o')
    pyautogui.keyUp('ctrl')
    wait_for_window('Open')
    
    print('\nLocate and open final plate of run.\n')
    # ensure smp is selected
    wait_for_window(smp, timeout=60)
    time.sleep(1)
    pyautogui.press('enter')
    progress()
    
    x = 1
    while x < file_amt:
        p = x
        print(f'Opening file {file_amt-x}')
        # open file explorer
        pyautogui.hotkey('ctrlleft', 'o')
        time.sleep(0.5)
        wait_for_window('Open')
        time.sleep(0.5)
    
        # navigate to data
        pyautogui.press('tab', presses=9)
        pyautogui.press('end')
    
        # select and open file
        if x <= file_amt:
            pyautogui.press('up', presses=x)
            print(f'up + {x+1}')
            time.sleep(0.1)
            pyautogui.press('enter')
            wait_for_window('SoftMax Pro')
            time.sleep(3)
            pyautogui.press('enter')
            time.sleep(1)
            print(f'Calling progress() with x = {x}, p = {p}')
            progress()
        x = x + 1
        print(f'x incremented to {x}')

    print(f'\nDone opening {x} files.\n')
   
    if file_amt == plate_amt:
        return
    else:
        messagebox.showinfo('Info', 'Please close the unnecessary files.')
        # required_keyword = "ok"
        # print(f'\nClose the unnecessary files and then type "ok" in the Terminal to continue with the script.\n')
        # user_input = input(f'Please enter "ok" to continue: ')
        # if user_input.lower() == required_keyword:
        #     print('Keyword accepted. Continuing with the script...')
        #     window_focus(smp)
        # else:
        #     print("Incorrect keyword. Script terminated.")
        #     sys.exit()

if __name__ == '__main__':
    main()