import os
import getpass
import json
from datetime import datetime
import os
import logging

# Ensure the logs directory exists
user = os.getlogin()
log_directory = os.path.join('progresses', 'debugging')
if not os.path.exists(log_directory):
    os.makedirs(log_directory)

# Logging setup
log_file_path = os.path.join(log_directory, f'{user}_log.log')
logging.basicConfig(filename=log_file_path, level=logging.DEBUG, 
                    format='%(asctime)s %(levelname)s:%(message)s')
script_name = os.path.basename(__file__)
logging.info("\n" + "="*50 + f"\nRunning {script_name}\n" + "="*50)

class UserCounter:
    def __init__(self):
        self.user = os.getlogin()
        with open(f'configs/config{self.user}.json', 'r') as file:
            self.config = json.load(file)
        self.plate_amt = None
        try:
            self.plate_amt = self.config['plate_amt']
        except KeyError:
            self.plate_amt = 0
        self.task = self.config['task']
        logging.info(self.task)
        self.user = getpass.getuser()
        self.count_file = os.path.join('progresses', 'logs', f'{self.task}_count.txt')
        self.times = self.get_times()

    def get_times(self):
        times = []
        try:
            with open(self.count_file, 'r') as file:
                data = file.readlines()
                for line in data:
                    parts = line.strip().split(':')
                    if len(parts) == 3:
                        user, times_str, plate_amt = parts
                        if user == self.user:
                            times = times_str.split(',')
                            return times
        except FileNotFoundError:
            return []
        return times

    def add_login_time(self):
        self.times.append(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        self.update_times_file()

    def update_times_file(self):
        updated = False
        try:
            with open(self.count_file, 'r') as file:
                data = file.readlines()
        except FileNotFoundError:
            data = []

        with open(self.count_file, 'w') as file:
            for line in data:
                parts = line.strip().split(':')
                if len(parts) == 3:
                    user, times_str, plate_amt = parts
                    if user == self.user:
                        times_str = ','.join(self.times)
                        file.write(f"{self.user} {times_str} {self.plate_amt}\n")
                        updated = True
                    else:
                        file.write(line)
                else:
                    file.write(line)
            if not updated:
                times_str = ','.join(self.times)
                file.write(f"{self.user} {times_str} {self.plate_amt}\n")

if __name__ == "__main__":
    counter = UserCounter()
    counter.add_login_time()
    logging.info(f"{', '.join(counter.times)}")