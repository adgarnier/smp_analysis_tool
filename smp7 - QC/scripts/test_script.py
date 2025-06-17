import time
import os
import sys
import subprocess
import pyautogui
import win32gui
import json
import logging
from tkinter import simpledialog

# plate_name = simpledialog.askstring("Input", "Enter your plate name:", )
plate_name = "PL-20250305-0001"
print(f'this is plate_name: {plate_name}')
start = plate_name[:-4]
last_4 = plate_name[-4:]
print(start)
print(last_4)



# split = plate_name.split('-')
# split_front = split[:2]
# split_back = split[2]
# print(f'this is split: {split}')
# print(f'this is split3: {split[2]}')
# print(f'this is split_front: {split_front}')
# print(f'this is split_back: {split_back}')

