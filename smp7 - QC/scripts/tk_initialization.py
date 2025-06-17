import sys
import subprocess
import logging
import os
import json
from tkinter import messagebox

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

plate_amt = 2
progress_data = {'current': 0, 'total': plate_amt, 'status': 'in-progress'}
x = 0

def main():
    check_and_install_pyautogui()
    check_and_install_pywin32()
    progress_data['status'] = 'done'
    with open(f'progresses/progress{user}.json', 'w') as progress_file:
        json.dump(progress_data, progress_file)
    logging.info('Commands complete.')
    messagebox.showinfo("Success", "Please restart SMP Analysis Tool.")

def progress():
    global x
    progress_data['current'] = x + 1
    with open(f'progresses/progress{user}.json', 'w') as progress_file:
        json.dump(progress_data, progress_file)
    x = x + 1

def check_and_install_pyautogui():
    try:
        import pyautogui
        logging.info('pyautogui is already installed.')
    except ImportError:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyautogui"])
            logging.info("pyautogui has been successfully installed.")
            messagebox.showinfo("Installation Success", "pyautogui has been successfully installed.")
        except subprocess.CalledProcessError as e:
            logging.error(f"Error installing pyautogui: {e}")
            messagebox.showerror("Installation Error", "Failed to install pyautogui.")
    progress()

def check_and_install_pywin32():
    try:
        import win32gui
        logging.info('pywin32 is already installed.')
    except ImportError:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
            logging.info("pywin32 has been successfully installed.")
            messagebox.showinfo("Installation Success", "pywin32 has been successfully installed.")
        except subprocess.CalledProcessError as e:
            logging.error(f"Error installing pywin32: {e}")
            messagebox.showerror("Installation Error", "Failed to install pywin32.")
    progress()

if __name__ == '__main__':
    main()