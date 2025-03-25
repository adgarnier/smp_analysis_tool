import subprocess
import logging
import sys
import os
import time

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Ensure the logs directory exists
user = os.getlogin()
log_directory = resource_path(os.path.join('progresses', 'debugging'))
if not os.path.exists(log_directory):
    os.makedirs(log_directory)

# Logging setup
log_file_path = os.path.join(log_directory, f'{user}_log.log')
logging.basicConfig(filename=log_file_path, level=logging.DEBUG, 
                    format='%(asctime)s %(levelname)s:%(message)s')
script_name = os.path.basename(__file__)
logging.info("\n" + "="*50 + f"\nRunning {script_name}\n" + "="*50)

def run_excel_extract_tool():
    time.sleep(1)
    script_path = resource_path(os.path.join('scripts', 'extract_GUI_2.1.1.py'))
    
    if not os.path.exists(script_path):
        logging.error("The extract_GUI script was not found in the 'scripts' directory. Ensure it is in the correct directory.")
        sys.exit(1)
    
    try:
        # logging.info('start of try block')
        logging.info("Starting Excel Extract Tool...")
        subprocess.run(["python", script_path], check=True)
        logging.info("Excel Extract Tool finished successfully.")
        # logging.info('end of try block')
    except subprocess.CalledProcessError as e:
        logging.error(f"An error occurred while running the Excel Extract Tool: {e}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    # logging.info('Running main')
    run_excel_extract_tool()