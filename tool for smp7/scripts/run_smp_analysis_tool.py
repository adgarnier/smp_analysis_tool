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

def run_smp_analysis_tool():
    # logging.info('Doing run_smp_analysis_tool')
    time.sleep(1)
    script_path = resource_path(os.path.join('scripts', 'smp_analysis_tool_v7.1.2.pyw'))
    
    if not os.path.exists(script_path):
        logging.error("The smp_analysis_tool.py script was not found in the 'scripts' directory. Ensure it is in the correct directory.")
        sys.exit(1)
    
    try:
        # logging.info('start of try block')
        logging.info("Starting SMP Analysis Tool...")
        subprocess.run(["pythonw", script_path], check=True)
        logging.info("SMP Analysis Tool finished successfully.")
        # logging.info('end of try block')
    except subprocess.CalledProcessError as e:
        logging.error(f"An error occurred while running the SMP Analysis Tool: {e}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    # logging.info('Running main')
    run_smp_analysis_tool()