import tkinter as tk
from tkinter import filedialog, messagebox, ttk, PhotoImage
import subprocess
import threading
import time
import json
import os
import sys
try:
    import pyautogui
except:
    print("pyautogui not installed")

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.create_widgets()
        self.process = None
        self.plate_amt = None
        self.file_amt = None
        self.directory = None
        self.setting_num = None
        self.x_cord = None
        self.y_cord = None
        self.export_setting = None
        self.finish_event = threading.Event()

        self.user = os.getlogin()

        self.buttons = [self.start_button, self.open_button, self.export_button, self.print_button, \
                   self.cycle_button, self.save_button, self.mask_button, self.full_button]
    
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

        tk.Label(frame_col1_top, text="# of files:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.entry_files = tk.Entry(frame_col1_top, width=5)
        self.entry_files.grid(row=1, column=1, padx=10, pady=5, sticky="w")

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
        self.listbox2.insert(tk.END, "Export")
        self.listbox2.insert(tk.END, "Print")

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

    def select_directory(self):
        global selected_directory
        # Open the directory selection dialog
        self.selected_directory = filedialog.askdirectory()
        # Update the label with the selected directory path
        self.dir_label.config(text="Folder ✅", fg="green")
    
    def variables(self):
        try:
            self.plate_amt = int(self.entry_plates.get())
        except ValueError:
            print("plate_amt is not set or invalid")
    
        try:
            self.file_amt = int(self.entry_files.get())
        except ValueError:
            print("file_amt is not set or invalid")

        if hasattr(self, 'selected_directory'):
            directory_value = self.selected_directory
            if directory_value is not None:
                self.directory = str(directory_value)
            else:
                print("directory is not set or invalid")
        else:
            print("selected_directory attribute is missing")

        try:
            self.setting_num = int(self.entry_setting.get())
        except ValueError:
            print("setting_num is not set or invalid")

        try:
            self.x_cord = int(self.entry_x.get())
        except ValueError:
            print("x_cord is not set or invalid")

        try:
            self.y_cord = int(self.entry_y.get())
        except ValueError:
            print("y_cord is not set or invalid")

        try:
            self.export_setting = self.listbox1.curselection()
        except ValueError:
            print("export is not set or invalid")
        
        # Create config_data using instance attributes
        variables = ['plate_amt', 'file_amt', 'directory', 'setting_num', 'x_cord', 'y_cord', 'export_setting']
        config_data = {}
        
        for var_name in variables:
            value = getattr(self, var_name)  # Use getattr to access instance attributes
            if value is not None:
                print(f"{var_name} is set to {value}")
                config_data[var_name] = value
            else:
                print(f"{var_name} is not set")
        
        # Write the collected data to the config.json file
        with open(f'configs/config{self.user}.json', 'w') as file:
            json.dump(config_data, file, indent=4)

        # init with the needed libraries
        try:
            import pyautogui
            import win32gui
        except ImportError:
            messagebox.showinfo("Info", "Click '🟢' in the top left corner to install the required programs.")
        
    # choosing python scipts
    def start(self):
        self.task = 'Start'
        self.variables()
        self._run_script("tk_01_start.py")
    
    def opening(self):
        self.task = 'Open'
        self.variables()
        if type(self.plate_amt) == int and type(self.file_amt) == int:
            print('Running open...')
            self._run_script("tk_02_open.py")
        elif type(self.plate_amt) != int and type(self.file_amt) != int:
            print('open plate_amt and file_amt not set')
            messagebox.showinfo("Info", "Please input a valid number of plates and files.")
        elif type(self.plate_amt) != int:
            print('open plate_amt not set')
            messagebox.showinfo("Info", "Please input a valid number of plates.")
        elif type(self.y_cord) != int:
            print('open file_amt not set')
            messagebox.showinfo("Info", "Please input a valid number of files.")
        else:
            print('Error in open')

    def export(self):
        self.task = 'Export'
        self.variables()
        selected_indices = self.listbox1.curselection()
        selected_items = [self.listbox1.get(i) for i in selected_indices]
        if not selected_indices:
            print('Nothing selected.')
            messagebox.showinfo("Info", "Please select export option(s).")
        elif type(self.plate_amt) == int:
            print('Running export...')
            # 0 = excel, 1 = lims, 2 = pnps
            if selected_indices == (0, 1, 2) or selected_indices == (0, 2):
                print('Invalid selection.')
                messagebox.showinfo("Info", "Invalid selection combination.")
            else:
                try: 
                    if 0 in selected_indices:
                        if not self.directory or self.directory == None:
                            messagebox.showinfo("Info", "Please select save directory.")
                        elif self.directory:
                            print(f'doing excel from {selected_items}')
                            self._run_script("tk_03_export.py")                            
                        else:
                            messagebox.showinfo("Info", "Please select save directory.")
                except ValueError:
                    print('oops')
                try:
                    if 1 in selected_indices:
                        print(f'doing lims from {selected_items}')
                        self._run_script("tk_03_export.py")
                except ValueError:
                    print('oops')
                try:
                    if 2 in selected_indices:
                        # messagebox.showinfo('Info', 'Still under development.')
                        if not self.directory or self.directory == None:
                            messagebox.showinfo("Info", "Please select save directory.")
                        elif self.directory:
                            print(f'doing pnps from {selected_items}')
                            self._run_script("tk_03_export.py")                            
                        else:
                            messagebox.showinfo("Info", "Please select save directory.")
                except ValueError:
                    print('oops')
        elif type(self.plate_amt) != int:
            print('export plate_amt not set')
            messagebox.showinfo("Info", "Please input a valid number of plates.")        
        else:
            print('Error in export')
    
    def printing(self):
        self.task = 'Print'
        self.variables()
        if type(self.plate_amt) == int:
            if self.checkbox_var3.get() == 1:
                print('Running print...')
                self.plate_amt = int(self.entry_plates.get())
                self.setting_num = 86
                variables = ['plate_amt', 'setting_num']
                config_data = {}
                for var_name in variables:
                    value = getattr(self, var_name)  # Use getattr to access instance attributes
                    if value is not None:
                        print(f"{var_name} is set to {value}")
                        config_data[var_name] = value
                    else:
                        print(f"{var_name} is not set")
                with open(f'configs/config{self.user}.json', 'w') as file:
                    json.dump(config_data, file, indent=4)
                self._run_script("tk_04_print.py")                
            elif type(self.plate_amt) == int and type(self.setting_num) == int:
                print('Running print...')
                self._run_script("tk_04_print.py")
            elif type(self.plate_amt) != int and type(self.setting_num) != int:
                print('print plate_amt and setting_amt not set')
                messagebox.showinfo("Info", "Please input a valid number of plates and setting #.")
            elif type(self.setting_num) != int:
                print('print setting_num not set')
                messagebox.showinfo("Info", "Please input a valid setting #.")
            else:
                print('Error in print')
        elif type(self.plate_amt) != int:
            print('print plate_amt not set')
            messagebox.showinfo("Info", "Please input a valid number of plates.")        
        else:
            print('Error in print')

    def cycle(self):
        self.task = 'Cycle'
        self.variables()
        if type(self.plate_amt) == int:
            print('Running cycle...')
            self._run_script("tk_05_cycle.py")
        elif type(self.plate_amt) != int:
            print('print plate_amt not set')
            messagebox.showinfo("Info", "Please input a valid number of plates.")        
        else:
            print('Error in cycle')
    
    def save(self):
        self.task = 'Save As Mod'
        self.variables()
        if type(self.plate_amt) == int:
            print('Running save...')
            # Check if the checkbox is selected (value 1 means checked)
            if self.checkbox_var1.get() == 1:
                print("Check box checked, continuing...")  # Perform the intended action here
                self._run_script("tk_06_saveasmod.py")
            elif self.checkbox_var1.get() == 0:
                print("Please check the 'Mods complete' checkbox before saving.")
                messagebox.showinfo("Info", "Please check the 'Mods complete' checkbox before saving.")
            else:
                print('Error in save')
        elif type(self.plate_amt) != int:
            print('print plate_amt not set')
            messagebox.showinfo("Info", "Please input a valid number of plates.")        
        else:
            print('Error in save')
    
    def mask(self):
        self.task = 'Masking'
        self.variables()
        if type(self.plate_amt) == int:
            if type(self.x_cord) != int or type(self.y_cord) != int:
                messagebox.showinfo("Info", "Move mouse to masking location and press enter.")
                messagebox.showinfo("Info", f"Enter the following coordinates:\n   x: {pyautogui.position().x}\n   y: {pyautogui.position().y}")
            elif self.checkbox_var2.get() == 0:
                pyautogui.moveTo(self.x_cord, self.y_cord)
                messagebox.showinfo("Info", "Is the mouse in the right position?\nConfirm by checking the 'Confirm' checkbox.")
                print('the end')
            elif self.checkbox_var2.get() == 1:
                print("Running mask...")  # Perform the intended action here
                self._run_script("tk_07_masking.py")
            else:
                print("Error in mask")
            # self._run_script("test.py")
        elif type(self.plate_amt) != int:
            print('print plate_amt not set')
            messagebox.showinfo("Info", "Please input a valid number of plates.")        
        else:
            print('Error in mask')
    
    def full(self):
        self.task = 'Full'
        self.variables()
        messagebox.showinfo('Info', 'Still under development.')
        return
        selected_indices = self.listbox2.curselection()
        selected_items = [self.listbox2.get(i) for i in selected_indices]
        if not selected_indices:
            print('Nothing selected.')
            messagebox.showinfo("Info", "Please select export option(s).")
        elif type(self.plate_amt) == int:
            print('Running full...')
            # 0 = start, 1 = open, 2 = export, 3 = print
            if selected_indices == (100):
                print('Invalid selection.')
                messagebox.showinfo("Info", "Invalid selection combination.")
            else:
                try: 
                    if 0 in selected_indices:
                        print('full start')
                        self.start()
                        self.wait_for_finish()
                except ValueError:
                    print('oops')
                try:
                    if 1 in selected_indices:
                        print('full open')
                        self.opening()
                        self.wait_for_finish()
                except ValueError:
                    print('oops')
                try:
                    if 2 in selected_indices:
                        print('full export')
                        self.export()
                except ValueError:
                    print('oops')
                try:
                    if 3 in selected_indices:
                        print('full print')
                        self.printing()
                except ValueError:
                    print('oops')
        elif type(self.plate_amt) != int:
            print('export plate_amt not set')
            messagebox.showinfo("Info", "Please input a valid number of plates.")        
        else:
            print('Error in export')

    def wait_for_finish(self): 
        self.finish_event.clear()
        self.finish_event.wait()
    
    def init_script(self):
        self._run_script("tk_initialization.py")
    
    def _run_script(self, script_name):
        # global running_script
        self.running_script = f'scripts/{script_name}'
        self.script()    
        
    # scripts
    def script(self):
        self.script_aborted = False
        self.progress.config(value=0)
        self.progress.config(maximum=100)
        # Disable the start button to prevent multiple instances running
        for button in self.buttons:
            button.config(state="disabled")
        # Run the "test.py" script in a separate thread to keep the GUI responsive
        threading.Thread(target=self._script_thread, daemon=True).start()

    def _script_thread(self):
        self.process = subprocess.Popen(["python", self.running_script])
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

        if self.process.poll() is None:
            self.master.after(500, self.check_progress)  # Continue checking progress if the process is still running
        else:
            self.finish_script()

    def finish_script(self):
        # Re-enable the run button
        for button in self.buttons:
            button.config(state="normal")
        # Show a pop-up message when the script ends if it was not aborted
        if not self.script_aborted:
            messagebox.showinfo("Script Status", f"{self.task} completed sucessfully.")
        self.progress["value"] = 0
        self.finish_event.set()
     
    def select_directory(self):
        # Open the directory selection dialog
        self.selected_directory = filedialog.askdirectory()
        # Update the label with the selected directory path
        if self.selected_directory:
            self.dir_label.config(text="Directory ✅", fg="green")
        else:
            messagebox.showinfo("Info", "Please select a save")

    def help_script(self):
        messagebox.showinfo("Info", "\
Start: Starts SoftMax Pro and initializes the software.\n\n\
Open: Opens the plates from the desired folder\n\
* When prompted, open the final file from the desired folder.\n\
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
Full: Under development.")
        
    def abort_script(self):
        try:
            import scripts\tk_00_base.py
            tk_00_base.timeout_abort()
            sys.exit()
        except SystemExit: 
            print("Exiting the script.")
            sys.exit()
        # Terminate the script immediately if it is running
        if hasattr(self, 'process') and self.process is not None:
            self.process.terminate()
            self.process = None
            self.script_aborted = True  # Set the abort flag
            messagebox.showinfo("Script Status", "Execution aborted.")
            # Re-enable the start button
            for button in self.buttons:
                button.config(state="normal") 
            # Stop updating the progress bar
            self.progress["value"] = 0

root = tk.Tk()
root.title("SMP Analysis Tool")
app = Application(master=root)
app.pack(fill="both", expand=True)
root.mainloop()
