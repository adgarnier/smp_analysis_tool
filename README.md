# 🧪 SMP Analysis Tool

## 🌟 Overview

**SMP Analysis Tool** is a graphical user interface (GUI) application built with Tkinter for automating various tasks related to plate data analysis using SoftMax Pro (SMP). 
It includes functionalities like starting the analysis, opening files, exporting data, signing plates, cycling through plates, saving data, masking wells, and running a sequence of tasks automatically.
There are versions available for SMP 6.5.2 and SMP 7.1.2

## 🛠 Prerequisites

- Python 3.x
- Tkinter
- pyautogui
- pywin32
- Other standard libraries (`os`, `json`, `logging`, `threading`, `subprocess`, `time`, `importlib`)

## 📥 Installation

1. **Clone or download the repository**:
   git clone https://github.com/yourusername/smp-analysis-tool.git cd smp-analysis-tool or
   Download ZIP, unzip.

3. **Install the required packages**:
   pip install pyautogui pywin32

4. **Ensure the necessary directories and files are in place**:

    - `images/IMG_4666.png`
    - `themes/azure/azure.tcl`
    - `scripts/*.pyw`
    - `configs/`
    - `progresses/`
    - `log_directory/`

## 🚀 Usage

1. **Run the executable**

2. **Interface**:

    - **Number of plates**: Enter the number of plates to be processed.
    - **Buttons**:
        - **Start**: Initialize the software.
        - **Open**: Open the specified number of plates.
        - **Export**: Export plate data.
        - **Sign**: Sign and submit plates.
        - **Print**: Prints the plates.
        - **Cycle**: Cycle through plates.
        - **Save**: Save plate data.
        - **Mask**: Mask specified wells.
        - **Full**: Run a sequence of tasks.
        - **Abort**: Abort the running task.
        - **Help**: Display information about tool usage.
    - **Select Folder**: Select the folder where exported data will be saved.
    - **Progress Bar**: Displays the progress of the running task.

<img src="https://github.com/user-attachments/assets/3d11e839-d98d-4a5a-9538-109f872bc51a" width="600">

## ⚙️ Configuration

The tool generates a configuration file for each user in the `configs/` directory. This file is used to store user-specific settings.

## 📝 Logging

Logs are saved in the `log_directory/` directory and provide detailed information about the tool's operations and any errors encountered.

## 🙏 Acknowledgements

- [SoftMax Pro](https://www.moleculardevices.com/products/microplate-readers/acquisition-and-analysis-software/softmax-pro-software) for inspiration.
- [Azure Theme for Tkinter](https://github.com/rdbende/Azure-ttk-theme) for the GUI theme.
