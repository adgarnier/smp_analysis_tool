import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import sys

class ScriptRunner:
    def __init__(self, root):
        self.root = root
        self.root.title("Script Runner")
        
        frame = tk.Frame(self.root)
        frame.pack(pady=10)
        
        start_label = tk.Label(frame, text="Start Value:")
        start_label.pack(side=tk.LEFT, padx=5)
        
        self.start_entry = tk.Entry(frame)
        self.start_entry.pack(side=tk.LEFT, padx=5)
        
        btn_gmcs = tk.Button(frame, text="Run extract_GMCs_new.py", command=lambda: self.run_script('extract_GMCs_new.py'))
        btn_gmcs.pack(side=tk.LEFT, padx=5)
        
        btn_ods = tk.Button(frame, text="Run extract_ODs_new.py", command=lambda: self.run_script('extract_ODs_new.py'))
        btn_ods.pack(side=tk.LEFT, padx=5)
        
        self.output_text = tk.Text(self.root, wrap=tk.WORD, height=15, width=80)
        self.output_text.pack(pady=10)
    
    def run_script(self, script_name):
        input_folder = filedialog.askdirectory(title="Select Input Folder")
        if not input_folder:
            messagebox.showerror("Error", "No folder selected.")
            return
        
        try:
            # Debugging information
            debug_info = (
                f"Running script: {script_name}\n"
                f"Input folder: {input_folder}\n"
                f"Current working directory: {os.getcwd()}\n"
                f"Python executable: {sys.executable}\n"
            )

            # Ensure the script uses the current directory
            script_path = os.path.join(os.getcwd(), script_name)

            # Run the script
            result = subprocess.run(
                ["python", script_path, input_folder],
                capture_output=True
            )

            # Clear the Text widget before inserting new content
            self.output_text.delete(1.0, tk.END)

            # Prepare the output to be inserted and copied
            output = debug_info + result.stdout.decode() + result.stderr.decode()

            # Insert debugging info and script output into the Text widget
            self.output_text.insert(tk.END, output)
            
            # Copy the output to clipboard
            self.root.clipboard_clear()
            self.root.clipboard_append(output)
            
            if result.returncode != 0:
                messagebox.showerror("Error", result.stderr.decode())
            else:
                messagebox.showinfo("Success", "Script executed successfully. Output copied to clipboard.")
        except Exception as e:
            # Clear the Text widget before inserting new content
            self.output_text.delete(1.0, tk.END)
            self.output_text.insert(tk.END, str(e))
            messagebox.showerror("Error", str(e))

root = tk.Tk()
app = ScriptRunner(root)
root.mainloop()