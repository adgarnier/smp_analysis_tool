import tkinter as tk
from tkinter import filedialog, messagebox, ttk, PhotoImage
import subprocess
import threading
import time
import json
import os
import logging
import importlib.util

try:
    import pyautogui
except:
    logging.info("pyautogui not installed")
try:
    import win32gui
except:
    logging.info("win32gui not installed")

tool_version = 'SMP Analysis Tool v3.3.2'

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

task = 'logon'
script_path = os.path.join('scripts', 'tk_user_logs.pyw')
subprocess.run(['pythonw', script_path])

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.icon_image()
        self.create_widgets()
        self.process = None
        self.plate_amt = None
        self.file_amt = None
        self.directory = None
        self.directory_open = None
        self.setting_num = None
        self.x_cord = None
        self.y_cord = None
        self.export_setting = None
        self.finish_event = threading.Event()
        self.script_aborted = None

        self.user = os.getlogin()

        self.buttons = [self.start_button, self.open_button, self.export_button, self.print_button, \
                   self.cycle_button, self.save_button, self.mask_button, self.full_button]
    
    def icon_image(self):
        self.image_path = r'images\IMG_4666.png'
        self.p1 = PhotoImage(file=self.image_path) 
        self.master.iconphoto(False, self.p1)

    def create_widgets(self):
        # Define button width
        button_width = 15
        
        # Number of plates
        self.label_plates = tk.Label(self, text="Number of plates:")
        self.entry_plates = tk.Entry(self)
        self.label_plates.grid(row=0, column=1, padx=10, pady=10, sticky="e")
        self.entry_plates.grid(row=0, column=2, padx=10, pady=10, sticky="w")
        
        # Two frames per column
        frame_col0_top = tk.Frame(self, borderwidth=2, relief="groove")
        frame_col1_top = tk.Frame(self, borderwidth=2, relief="groove")
        frame_col2_top = tk.Frame(self, borderwidth=2, relief="groove")
        frame_col3_top = tk.Frame(self, borderwidth=2, relief="groove")
        frame_col0_bottom = tk.Frame(self, borderwidth=2, relief="groove")
        frame_col1_bottom = tk.Frame(self, borderwidth=2, relief="groove")
        frame_col2_bottom = tk.Frame(self, borderwidth=2, relief="groove")
        frame_col3_bottom = tk.Frame(self, borderwidth=2, relief="groove")
    
        # Position the frames
        frame_col0_top.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        frame_col1_top.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
        frame_col2_top.grid(row=1, column=2, padx=5, pady=5, sticky="nsew")
        frame_col3_top.grid(row=1, column=3, padx=5, pady=5, sticky="nsew")
        frame_col0_bottom.grid(row=2, column=0, padx=5, pady=5, sticky="nsew")
        frame_col1_bottom.grid(row=2, column=1, padx=5, pady=5, sticky="nsew")
        frame_col2_bottom.grid(row=2, column=2, padx=5, pady=5, sticky="nsew")
        frame_col3_bottom.grid(row=2, column=3, padx=5, pady=5, sticky="nsew")
    
        # Add weight to columns to make the buttons expand with the window

        for i in range(4):
            self.columnconfigure(i, weight=1)
        for i in range(4):
            self.rowconfigure(i, weight=1)

        frames = [frame_col0_top, frame_col1_top, frame_col2_top, frame_col3_top, \
                  frame_col0_bottom, frame_col1_bottom, frame_col2_bottom, frame_col3_bottom]

        for frame in frames:
            for i in range(2):
                frame.columnconfigure(i, weight=1)

        # Widgets inside the top frames
        # Start section inside frame_col0_top
        self.start_button = tk.Button(frame_col0_top, text="Start", command=self.start, width=button_width)
        self.start_button.pack(padx=10, pady=10, fill="x")
        
        # Open section inside frame_col1_top
        self.open_button = tk.Button(frame_col1_top, text="Open", command=self.opening, width=button_width)
        self.open_button.grid(row=0, column=0, padx=10, pady=10, columnspan=2)

        self.dir_label_open = tk.Label(frame_col1_top, text="No folder selected")
        self.dir_label_open.grid(row=1, column=0, padx=10, pady=5, columnspan=2)
        select_dir_button_open = tk.Button(frame_col1_top, text="Select Folder", command=self.select_directory_open)
        select_dir_button_open.grid(row=2, column=0, padx=10, pady=5, columnspan=2)

        tk.Label(frame_col1_top, text="# of files:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.entry_files = tk.Entry(frame_col1_top, width=5)
        self.entry_files.grid(row=3, column=1, padx=10, pady=5, sticky="w")

        # Export section inside frame_col2_top
        self.export_button = tk.Button(frame_col2_top, text="Export", command=self.export, width=button_width)
        self.export_button.pack(padx=10, pady=10, fill="x")
        
        self.dir_label = tk.Label(frame_col2_top, text="No folder selected")
        self.dir_label.pack(padx=10, pady=5)
        select_dir_button = tk.Button(frame_col2_top, text="Select Folder", command=self.select_directory)
        select_dir_button.pack(padx=10, pady=5)

        listbox1_frame = tk.Frame(frame_col2_top) 
        listbox1_frame.pack(padx=10, pady=10, fill="both", expand=True)

        self.listbox1 = tk.Listbox(listbox1_frame, selectmode="single", height=3, width=1)
        self.listbox1.pack(side="left", fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(listbox1_frame, orient="vertical", command=self.listbox1.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox1.config(yscrollcommand=scrollbar.set)

        self.listbox1.insert(tk.END, "Excel")
        self.listbox1.insert(tk.END, "LIMS")
        self.listbox1.insert(tk.END, "PnPS")
        
        # Print section inside frame_col3_top
        self.print_button = tk.Button(frame_col3_top, text="Print", command=self.printing, width=button_width)
        self.print_button.grid(row=0, column=0, padx=10, pady=10, columnspan=2)

        tk.Label(frame_col3_top, text="Setting #:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_setting = tk.Entry(frame_col3_top, width=5)
        self.entry_setting.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        self.checkbox_var3 = tk.IntVar()
        save_checkbox = tk.Checkbutton(frame_col3_top, text="Skip", variable=self.checkbox_var3)
        save_checkbox.grid(row=2, column=0, padx=10, pady=5, columnspan=2)
        
        # Widgets inside the bottom frames
        # Cycle section inside frame_col0_bottom
        self.cycle_button = tk.Button(frame_col0_bottom, text="Cycle", command=self.cycle, width=button_width)
        self.cycle_button.pack(padx=10, pady=10, fill="x")
        
        # Save section inside frame_col2_bottom
        self.save_button = tk.Button(frame_col1_bottom, text="Save As Mod", command=self.save, width=button_width)
        self.save_button.pack(padx=10, pady=10, fill="x")
        
        self.checkbox_var1 = tk.IntVar()
        save_checkbox = tk.Checkbutton(frame_col1_bottom, text="Mods complete", variable=self.checkbox_var1)
        save_checkbox.pack(padx=10, pady=10)
    
        # Masking section inside frame_col2_bottom
        self.mask_button = tk.Button(frame_col2_bottom, text="Masking", command=self.mask, width=button_width)
        self.mask_button.grid(row=0, column=0, padx=10, pady=10, columnspan=2)
        
        tk.Label(frame_col2_bottom, text="x:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_x = tk.Entry(frame_col2_bottom, width=10)
        self.entry_x.grid(row=1, column=1, padx=10, pady=5, sticky="w")
    
        tk.Label(frame_col2_bottom, text="y:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.entry_y = tk.Entry(frame_col2_bottom, width=10)
        self.entry_y.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        self.checkbox_var2 = tk.IntVar()
        save_checkbox = tk.Checkbutton(frame_col2_bottom, text="Confirm", variable=self.checkbox_var2)
        save_checkbox.grid(row=3, column=0, padx=10, pady=5, columnspan=2)
        
        # Full section inside frame_col3_bottom    
        self.full_button = tk.Button(frame_col3_bottom, text="Full", command=self.full, width=button_width)
        self.full_button.pack(padx=10, pady=10, fill="x")

        listbox2_frame = tk.Frame(frame_col3_bottom)
        listbox2_frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.listbox2 = tk.Listbox(listbox2_frame, selectmode="multiple", height=3, width=3)
        self.listbox2.pack(side="left", fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(listbox2_frame, orient="vertical", command=self.listbox2.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox2.config(yscrollcommand=scrollbar.set)
        
        self.listbox2.insert(tk.END, "Start")
        self.listbox2.insert(tk.END, "Open")
        self.listbox2.insert(tk.END, "Masking")
        self.listbox2.insert(tk.END, "Save As Mod")
        self.listbox2.insert(tk.END, "Cycle")
        self.listbox2.insert(tk.END, "Print")
        self.listbox2.insert(tk.END, "Export LIMS")
        self.listbox2.insert(tk.END, "Export Excel")
        self.listbox2.insert(tk.END, "Export PnPS")

        # init button
        self.init_button = tk.Button(self, text="🟢", command=self.init_script, width=3)
        self.init_button.grid(row=0, column=0, padx=10, pady=10, sticky="w")  
        
        # Help button
        self.help_button = tk.Button(self, text="❔", command=self.help_script, width=3)
        self.help_button.grid(row=0, column=3, padx=10, pady=10, sticky="e")  
      
        # Abort button
        self.abort_button = tk.Button(self, text="Abort", command=self.abort_script, width=button_width)
        self.abort_button.grid(row=3, column=3, padx=10, pady=10, sticky="e")
        self.abort_button.config(bg="indianred1", fg="red4")
              
        # Progress bar
        self.progress = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

    def variables(self):
        try:
            self.plate_amt = int(self.entry_plates.get())
        except ValueError:
            logging.info("plate_amt is not set or invalid")
    
        try:
            self.file_amt = int(self.entry_files.get())
        except ValueError:
            logging.info("file_amt is not set or invalid")

        if hasattr(self, 'selected_directory'):
            directory_value = self.selected_directory
            if directory_value is not None:
                self.directory = str(directory_value)
                self.directory = self.directory.replace('/', '\\')
            else:
                logging.info("directory is not set or invalid")
        else:
            logging.info("selected_directory attribute is missing")

        if hasattr(self, 'selected_directory_open'):
            directory_value_open = self.selected_directory_open
            if directory_value_open is not None:
                self.directory_open = str(directory_value_open)
                self.directory_open = self.directory_open.replace('/', '\\')
            else:
                logging.info("directory is not set or invalid")
        else:
            logging.info("selected_directory attribute is missing")

        try:
            self.setting_num = int(self.entry_setting.get())
        except ValueError:
            logging.info("setting_num is not set or invalid")

        try:
            self.x_cord = int(self.entry_x.get())
        except ValueError:
            logging.info("x_cord is not set or invalid")

        try:
            self.y_cord = int(self.entry_y.get())
        except ValueError:
            logging.info("y_cord is not set or invalid")

        try:
            self.export_setting = self.listbox1.curselection()
        except ValueError:
            logging.info("export is not set or invalid")
        
        try:
            self.export_setting2 = self.listbox2.curselection()
        except ValueError:
            logging.info("export is not set or invalid")

        try:
            self.task = self.task
        except ValueError:
            logging.info("here")

        # Create config_data using instance attributes
        variables = ['plate_amt', 'file_amt', 'directory', 'directory_open', 'setting_num', 'x_cord', 'y_cord', 'export_setting', 'export_setting2', 'task']
        config_data = {}
        
        for var_name in variables:
            value = getattr(self, var_name)  # Use getattr to access instance attributes
            if value is not None:
                logging.info(f"{var_name} is set to {value}")
                config_data[var_name] = value
            else:
                logging.info(f"{var_name} is not set")
        
        # Write the collected data to the config.json file
        with open(f'configs/config{self.user}.json', 'w') as file:
            json.dump(config_data, file, indent=4)

        # Check if pyautogui and win32gui are installed
        pyautogui_installed = importlib.util.find_spec("pyautogui") is not None
        win32gui_installed = importlib.util.find_spec("win32gui") is not None

        if not pyautogui_installed or not win32gui_installed:
            messagebox.showinfo("Info", "Click '🟢' in the top left corner to install the required programs.")
            return
        else:
            logging.info("pyautogui and win32gui are installed")

    # choosing pythonw scipts
    def start(self):
        self.task = 'Start'
        self.variables()
        if importlib.util.find_spec("pyautogui") is not None:
            self._run_script("tk_01_start.pyw")
    
    def opening(self):
        self.task = 'Open'
        self.variables()
        if importlib.util.find_spec("pyautogui") is not None:
            if not self.directory_open or self.directory_open == None:
                messagebox.showinfo("Info", "Please select save directory.")
            elif type(self.plate_amt) == int and type(self.file_amt) == int:
                logging.info('Running open...')
                self._run_script("tk_02_open.pyw")
            elif type(self.plate_amt) != int and type(self.file_amt) != int:
                logging.info('open plate_amt and file_amt not set')
                messagebox.showinfo("Info", "Please input a valid number of plates and files.")
            elif type(self.plate_amt) != int:
                logging.info('open plate_amt not set')
                messagebox.showinfo("Info", "Please input a valid number of plates.")
            elif type(self.file_amt) != int:
                logging.info('open file_amt not set')
                messagebox.showinfo("Info", "Please input a valid number of files.")
            else:
                logging.info('Error in open')

    def export(self):
        self.task = 'Export'
        self.variables()
        selected_indices = self.listbox1.curselection()
        selected_items = [self.listbox1.get(i) for i in selected_indices]
        selected_indices2 = self.listbox2.curselection()
        selected_items2 = [self.listbox2.get(i) for i in selected_indices2]
        logging.info(selected_indices2)
        if importlib.util.find_spec("pyautogui") is not None:
            if not selected_indices and not selected_indices2:
                logging.info('Nothing selected.')
                messagebox.showinfo("Info", "Please select export option(s).")
            elif type(self.plate_amt) == int:
                logging.info('Running export...')
                # 0 = excel, 1 = lims, 2 = pnps
                if selected_indices == (0, 1, 2) or selected_indices == (0, 2):
                    logging.info('Invalid selection.')
                    messagebox.showinfo("Info", "Invalid selection combination.")
                else:
                    try: 
                        if 0 in selected_indices or 7 in selected_indices2:
                            if not self.directory or self.directory == None:
                                messagebox.showinfo("Info", "Please select save directory.")
                            elif self.directory:
                                logging.info(f'doing excel from {selected_items} or {selected_items2}')
                                self._run_script("tk_03_export.pyw")                            
                            else:
                                messagebox.showinfo("Info", "Please select save directory.")
                    except ValueError:
                        logging.info('oops')
                    try:
                        if 1 in selected_indices or 6 in selected_indices2:
                            logging.info(f'doing lims from {selected_items} or {selected_items2}')
                            self._run_script("tk_03_export.pyw")
                    except ValueError:
                        logging.info('oops')
                    try:
                        if 2 in selected_indices or 8 in selected_indices2:
                            if not self.directory or self.directory == None:
                                messagebox.showinfo("Info", "Please select save directory.")
                            elif self.directory:
                                logging.info(f'doing pnps from {selected_items} or {selected_items2}')
                                self._run_script("tk_03_export.pyw")                            
                            else:
                                messagebox.showinfo("Info", "Please select save directory.")
                    except ValueError:
                        logging.info('oops')
            elif type(self.plate_amt) != int:
                logging.info('export plate_amt not set')
                messagebox.showinfo("Info", "Please input a valid number of plates.")        
            else:
                logging.info('Error in export')
    
    def printing(self):
        self.task = 'Print'
        self.variables()
        if importlib.util.find_spec("pyautogui") is not None:
            if type(self.plate_amt) == int:
                if self.checkbox_var3.get() == 1:
                    logging.info('Running print...')
                    self.plate_amt = int(self.entry_plates.get())
                    self.setting_num = 86
                    variables = ['plate_amt', 'setting_num', 'task']
                    config_data = {}
                    for var_name in variables:
                        value = getattr(self, var_name)  # Use getattr to access instance attributes
                        if value is not None:
                            logging.info(f"{var_name} is set to {value}")
                            config_data[var_name] = value
                        else:
                            logging.info(f"{var_name} is not set")
                    with open(f'configs/config{self.user}.json', 'w') as file:
                        json.dump(config_data, file, indent=4)
                    self._run_script("tk_04_print.pyw")                
                elif type(self.plate_amt) == int and type(self.setting_num) == int:
                    logging.info('Running print...')
                    self._run_script("tk_04_print.pyw")
                elif type(self.plate_amt) != int and type(self.setting_num) != int:
                    logging.info('print plate_amt and setting_amt not set')
                    messagebox.showinfo("Info", "Please input a valid number of plates and setting #.")
                elif type(self.setting_num) != int:
                    logging.info('print setting_num not set')
                    messagebox.showinfo("Info", "Please input a valid setting #.")
                else:
                    logging.info('Error in print')
            elif type(self.plate_amt) != int:
                logging.info('print plate_amt not set')
                messagebox.showinfo("Info", "Please input a valid number of plates.")        
            else:
                logging.info('Error in print')

    def cycle(self):
        self.task = 'Cycle'
        self.variables()
        if importlib.util.find_spec("pyautogui") is not None:
            if type(self.plate_amt) == int:
                logging.info('Running cycle...')
                self._run_script("tk_05_cycle.pyw")
            elif type(self.plate_amt) != int:
                logging.info('print plate_amt not set')
                messagebox.showinfo("Info", "Please input a valid number of plates.")        
            else:
                logging.info('Error in cycle')
    
    def save(self):
        self.task = 'Save As Mod'
        self.variables()
        if importlib.util.find_spec("pyautogui") is not None:
            if type(self.plate_amt) == int:
                logging.info('Running save...')
                # Check if the checkbox is selected (value 1 means checked)
                if self.checkbox_var1.get() == 1:
                    logging.info("Check box checked, continuing...")  # Perform the intended action here
                    self._run_script("tk_06_saveasmod.pyw")
                elif self.checkbox_var1.get() == 0:
                    logging.info("Please check the 'Mods complete' checkbox before saving.")
                    messagebox.showinfo("Info", "Please check the 'Mods complete' checkbox before saving.")
                else:
                    logging.info('Error in save')
            elif type(self.plate_amt) != int:
                logging.info('print plate_amt not set')
                messagebox.showinfo("Info", "Please input a valid number of plates.")        
            else:
                logging.info('Error in save')
    
    def mask(self):
        self.task = 'Masking'
        self.variables()
        if importlib.util.find_spec("pyautogui") is not None:
            if type(self.plate_amt) == int:
                if type(self.x_cord) != int or type(self.y_cord) != int:
                    messagebox.showinfo("Info", "Move mouse to masking location and press enter.")
                    messagebox.showinfo("Info", f"Enter the following coordinates:\n   x: {pyautogui.position().x}\n   y: {pyautogui.position().y}")
                elif self.checkbox_var2.get() == 0:
                    pyautogui.moveTo(self.x_cord, self.y_cord)
                    messagebox.showinfo("Info", "Is the mouse in the right position?\nConfirm by checking the 'Confirm' checkbox.")
                    logging.info('the end')
                elif self.checkbox_var2.get() == 1:
                    logging.info("Running mask...")  # Perform the intended action here
                    self._run_script("tk_07_masking.pyw")
                else:
                    logging.info("Error in mask")
                # self._run_script("test.pyw")
            elif type(self.plate_amt) != int:
                logging.info('print plate_amt not set')
                messagebox.showinfo("Info", "Please input a valid number of plates.")        
            else:
                logging.info('Error in mask')

    def full(self):
        def run_full():
            messagebox.showinfo("Info", "Under development.")
            return

            self.task = 'Full'
            self.script_aborted = False
            self.variables()
            
            selected_indices2 = self.listbox2.curselection()
            selected_items = [self.listbox2.get(i) for i in selected_indices2]
            logging.info(", ".join(selected_items))
            
            if not selected_indices2:
                logging.info('Nothing selected.')
                messagebox.showinfo("Info", "Please select export option(s).")
                return
            
            if type(self.plate_amt) != int:
                logging.info('plate_amt not set')
                messagebox.showinfo("Info", "Please input a valid number of plates.")
                return
            
            if selected_indices2 == (10000):
                logging.info('Invalid selection.')
                messagebox.showinfo("Info", "Invalid selection combination.")
                return
            else:
                logging.info('checks')
                # check inputs are there
                try:
                    if 1 in selected_indices2:
                        if type(self.file_amt) != int:
                            logging.info('open file_amt not set')
                            messagebox.showinfo("Info", "Please input a valid number of files.")
                            return
                except ValueError:
                    logging.info('oops')
                try:
                    if 2 in selected_indices2:
                        if type(self.x_cord) != int or type(self.y_cord) != int:
                            messagebox.showinfo("Info", "Please input the correct x and y coordinates.")
                            return
                except ValueError:
                    logging.info('oops')
                try:
                    if 3 in selected_indices2:
                        if self.checkbox_var1.get() == 0:
                            logging.info("Please check the 'Mods complete' checkbox before saving.")
                            messagebox.showinfo("Info", "Please check the 'Mods complete' checkbox before saving.")
                            return
                except ValueError:
                    logging.info('oops')
                try:
                    if 5 in selected_indices2:
                        if type(self.setting_num) != int and self.checkbox_var3.get() != 1:
                            logging.info('print setting_num not set')
                            messagebox.showinfo("Info", "Please input a valid setting #.")
                            return
                except ValueError:
                    logging.info('oops')
                try:
                    if 7 in selected_indices2 or 8 in selected_indices2:
                        if not self.directory or self.directory == None:
                            messagebox.showinfo("Info", "Please select save directory.")
                            return
                except ValueError:
                    logging.info('oops')

                logging.info('Running full...')
                
                actions = {
                    0: 'start',
                    1: 'opening',
                    2: 'mask',
                    3: 'save',
                    4: 'cycle',
                    5: 'printing',
                    6: 'export', # lims
                    7: 'export', # excel
                    8: 'export' # pnps
                }
                
                for index in selected_indices2:
                    if index in actions and not self.script_aborted:
                        go = 1
                        action_method = getattr(self, actions[index], None)
                        if callable(action_method):
                            logging.info(f'Doing {actions[index]}...')
                            action_method()
                        time.sleep(1)
                        while go == 1:
                            try:
                                # logging.info('looking for abort status')
                                if self.find_window_by_title('Abort Status'):
                                    logging.info("Found 'Abort Status'")
                                    return
                                # logging.info('did not find abort')
                            except Exception as e:
                                logging.info(f'Error checking abort status: {e}')
                            try:
                                # logging.info('looking for script stat')
                                if self.find_window_by_title('Script Status'):
                                    logging.info("Found 'Scrip Status'")
                                    go = 0
                                # logging.info('tool not found')
                            except Exception as e:
                                logging.info(f'Error checking script status: {e}')
                            time.sleep(0.5)
                        pyautogui.press('space')
                
                if not self.script_aborted:
                    if len(selected_items) == 1:
                        messagebox.showinfo("Script Status", f"{selected_items[0]} completed.")
                    else:
                        messagebox.showinfo("Script Status", f"{', '.join(selected_items[:-1])} and {selected_items[-1]} completed.")
                else:
                    messagebox.showerror("Abort Status", "Execution aborted.")
                    return
        
        self.thread = threading.Thread(target=run_full)
        self.thread.start()
    
    def init_script(self):
        self.task = 'Initialization'
        self._run_script("tk_initialization.py")
    
    def _run_script(self, script_name):
        # global running_script
        self.running_script = f'scripts/{script_name}'
        log_path = os.path.join('scripts', 'tk_process_logs.pyw')
        subprocess.run(['pythonw', log_path])
        self.script()    
        
    # scripts

    def find_window_by_title(self, title):
        hwnd = win32gui.FindWindow(None, title)
        return hwnd

    def script(self):
        self.script_aborted = False
        self.progress.config(value=0)
        self.progress.config(maximum=100)
        # Disable the start button to prevent multiple instances running
        for button in self.buttons:
            button.config(state="disabled")
        # Run the "test.pyw" script in a separate thread to keep the GUI responsive
        threading.Thread(target=self._script_thread, daemon=True).start()

    def _script_thread(self):
        self.process = subprocess.Popen(["pythonw", self.running_script])
        self.master.after(500, self.check_progress)  # Start checking progress every 500ms

    def check_progress(self):
        if os.path.exists(f'progresses/progress{self.user}.json'):
            with open(f'progresses/progress{self.user}.json', 'r') as progress_file:
                progress_data = json.load(progress_file)

                if progress_data['status'] == 'done':
                    self.progress.config(value=100)
                    self.process.wait()
                    self.finish_script()
                    return

                current_progress = progress_data.get('current', 0)
                total_progress = progress_data.get('total', 100)

                # Update the progress bar values
                self.progress.config(maximum=total_progress)
                self.progress.config(value=current_progress)
                self.master.update_idletasks()

        if self.process.poll() is None and not self.find_window_by_title('Abort Status'):
            self.master.after(500, self.check_progress)  # Continue checking progress if the process is still running
        else:
            self.finish_script()

    def finish_script(self):
        # Re-enable the run button
        for button in self.buttons:
            button.config(state="normal")
        # Show a pop-up message when the script ends if it was not aborted
        if not self.script_aborted and not self.find_window_by_title('Abort Status'):
            messagebox.showinfo("Script Status", f"{self.task} completed sucessfully.")
        progress_data = {'current': 0, 'total': self.plate_amt, 'status': 'in-progress'}
        with open(f'progresses/progress{user}.json', 'w') as progress_file:
            json.dump(progress_data, progress_file)
        self.progress["value"] = 0
        self.finish_event.set()
     
    def select_directory(self):
        # Open the directory selection dialog
        self.selected_directory = filedialog.askdirectory()
        # Update the label with the selected directory path
        if self.selected_directory:
            self.dir_label.config(text="Folder ✅", fg="green")
        else:
            messagebox.showinfo("Info", "Please select a save")

    def select_directory_open(self):
        # Open the directory selection dialog
        self.selected_directory_open = filedialog.askdirectory()
        # Update the label with the selected directory path
        if self.selected_directory_open:
            self.dir_label_open.config(text="Folder ✅", fg="green")
        else:
            messagebox.showinfo("Info", "Please select a save")

    def help_script(self):
        messagebox.showinfo("Info", "\
Start: Starts SoftMax Pro and initializes the software.\n\n\
Open: Opens the plates from a selected folder.\n\
* Opens all the plates from a selected folder.\n\
** '# of files' is the amount of files in the folder (includes mods).\n\n\
Export: Exports the plate data.\n\
* 'Excel' exports all the data to an excel file in a selected folder.\n\
** 'LIMS' does not need a selected folder to export correctly. \n\n\
Print: Prints the plates. \n\
* 'Setting #' is how many settings your SMP preset 'Printings Setting' is under 'Load settings' in 'Printing Preferences' excluding '(NONE)').\n\
* 'Skip' skips the process of selecting 'Setting #'. \n\n\
Cycle: Cycles through the plates without impacting them. \n\n\
Save As Mod: Saves the plates as '<Data ##-##-##-#####_mod>'. \n\n\
Masking: Masks a specified well of the plates. \n\n\
Full: Selected tasks are completed sequencially.")

    def abort_script(self):
        # Terminate the script immediately if it is running
        if hasattr(self, 'process') and self.process is not None:
            self.process.terminate()
            self.process = None
            self.script_aborted = True  # Set the abort flag
            messagebox.showerror("Abort Status", "Execution aborted.")
            # Re-enable the start button
            for button in self.buttons:
                button.config(state="normal") 
            # Stop updating the progress bar
            self.progress["value"] = 0

root = tk.Tk()
root.title(tool_version)
app = Application(master=root)
app.pack(fill="both", expand=True)
root.mainloop()
