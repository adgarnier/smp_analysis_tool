import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import logging
import json

class ScriptRunner:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Extract Tool v2.1.1")
        self.user = os.getlogin()
        
        # Ensure the logs directory exists
        log_directory = os.path.join('progresses', 'debugging')
        if not os.path.exists(log_directory):
            os.makedirs(log_directory)

        # Logging setup
        log_file_path = os.path.join(log_directory, f'{self.user}_log.log')
        logging.basicConfig(filename=log_file_path, level=logging.DEBUG, 
                            format='%(asctime)s %(levelname)s:%(message)s')
        script_name = os.path.basename(__file__)
        logging.info("\n" + "="*50 + f"\nRunning {script_name}\n" + "="*50)

        # GUI
        frame = tk.Frame(self.root)
        frame.pack(pady=10)
        
        btn_gmcs = tk.Button(frame, text="       Extract GMCs       ", command=lambda: self.run_script('extract_GMCs_new.py'))
        btn_gmcs.pack(side=tk.LEFT, padx=5)
        
        btn_ods = tk.Button(frame, text="       Extract ODs       ", command=lambda: self.run_script('extract_ODs_new.py'))
        btn_ods.pack(side=tk.LEFT, padx=5)
        
        # Frame around Radio button group
        radio_frame_outer = tk.LabelFrame(self.root, text="Select Test", padx=10, pady=5)
        radio_frame_outer.pack(pady=5)
        
        self.radio_var = tk.StringVar(value="Default")
        radio_frame = tk.Frame(radio_frame_outer)
        radio_frame.pack(pady=5)
        
        radio1 = tk.Radiobutton(radio_frame, text="Default", variable=self.radio_var, value="Default")
        radio1.pack(side=tk.LEFT, padx=5)
        
        radio2 = tk.Radiobutton(radio_frame, text="VzV", variable=self.radio_var, value="VzV")
        radio2.pack(side=tk.LEFT, padx=5)
        
        radio3 = tk.Radiobutton(radio_frame, text="GSK ELLA", variable=self.radio_var, value="GSK ELLA")
        radio3.pack(side=tk.LEFT, padx=5)
        
        self.output_text = tk.Text(self.root, wrap=tk.WORD, height=10, width=75)
        self.output_text.pack(pady=10)

        self.clear_terminal()

        test_text = (
            "Extract GMCs and raw ODs from Excel files exported from SoftMax Pro\n\n"
            "Instructions:\n"
            "1. Click on either \'Extract GMCs\' or \'Extract ODs\'\n"
            "2. Select the folder that contains your exported excel files from SMP\n"
            "3. Copy the output in the \'Windows PowerShell\' window to a new Excel book\n"
            "4. Go to Data -> Click Text to Columns -> Select Delimited -> Click Next -> Checkmark Comma -> Click Finish\n"
            "5. Complete analysis in Excel 😊\n"
        )
        self.output_text.insert(tk.END, test_text)
    
    def clear_terminal(self):
        # Clear the terminal based on the operating system
        if os.name == 'nt':  # For Windows
            os.system('cls')

    def variables(self):
        try:
            self.select_test = self.radio_var.get()
        except ValueError:
            logging.info("select_test is not set or invalid")

        # Create config_data using instance attributes
        variables = ['select_test']
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

    def run_script(self, script_name):
        self.clear_terminal()
        self.variables()
        input_folder = filedialog.askdirectory(title="Select Input Folder")
        if not input_folder:
            messagebox.showerror("Error", "No folder selected.")
            return
        
        try:
            # Debugging information
            debug_info = (
                f"\nRunning script: {script_name}\n"
                f"Input folder: {input_folder}\n"
                f"Current working directory: {os.getcwd()}\n"
                f"Selected option: {self.radio_var.get()}\n"
                "Script complete\n"
                # f"Python executable: {sys.executable}\n"
            )

            # Ensure the script uses the current directory
            script_path = os.path.join(os.getcwd(), 'scripts', script_name)

            # Run the script
            result = subprocess.run(
                ["python", script_path, input_folder],
                capture_output=False
            )

            # Clear the Text widget before inserting new content
            # self.output_text.delete(1.0, tk.END)

            # Insert debugging info and script output into the Text widget
            self.output_text.insert(tk.END, debug_info)
            
            if result.returncode != 0:
                messagebox.showerror("Error", result.stderr)
            else:
                messagebox.showinfo("Success", "Script executed successfully!")
            
        except Exception as e:
            # Clear the Text widget before inserting new content
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, str(e))
            messagebox.showerror("Error", str(e))

root = tk.Tk()
app = ScriptRunner(root)
root.mainloop()