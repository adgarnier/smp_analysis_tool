import time
import os
import sys
import subprocess
import pyautogui
import win32gui
import json
from tk_00_base import *

def main():
    test()

def _run_script(self, script_name):
    # global running_script
    self.running_script = script_name
    self.script() 

if __name__ == '__main__':
    main()