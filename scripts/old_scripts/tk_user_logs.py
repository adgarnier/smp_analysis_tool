import os
import getpass
from datetime import datetime

class UserCounter:
    def __init__(self):
        self.user = getpass.getuser()
        self.count_file = os.path.join('progresses', 'logs', 'user_count.txt')
        self.times = self.get_times()

    def get_times(self):
        times = []
        try:
            with open(self.count_file, 'r') as file:
                data = file.readlines()
                for line in data:
                    parts = line.strip().split(':')
                    if len(parts) == 2:
                        user, times_str = parts
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
                if len(parts) == 2:
                    user, times_str = parts
                    if user == self.user:
                        times_str = ','.join(self.times)
                        file.write(f"{self.user}:{times_str}\n")
                        updated = True
                    else:
                        file.write(line)
                else:
                    file.write(line)
            if not updated:
                times_str = ','.join(self.times)
                file.write(f"{self.user} {times_str}\n")

if __name__ == "__main__":
    counter = UserCounter()
    counter.add_login_time()
    print(f"{', '.join(counter.times)}")
