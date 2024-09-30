
# XLS/CSV File Converter

This repository contains three Python scripts that convert Excel files (.xls, .xlsx) to CSV format. Each script serves different purposes, whether you prefer a graphical interface or command-line usage, with additional features like automated file transfer via SCP.

## xlsToCsvVisual.py

This script provides a graphical interface for converting Excel files to CSV.

### Requirements

Before running the application, ensure you have the following Python packages installed:
- `pandas`
- `tkinterdnd2`
- `python-magic`
- `tkinter`

### Install dependencies:

```bash
pip install pandas python-magic
```

To run the script, execute:

```bash
python xlsToCsvVisual.py
```

## xlsToCsvTerminal.py

This is a command-line version of the application that converts Excel files to CSV without the graphical interface.

### Requirements

Ensure the following packages are installed:
- `pandas`
- `python-magic`

### Install dependencies:

```bash
pip install pandas python-magic
```

To run the script, execute:

```bash
python xlsToCsvTerminal.py input_file.xls
```

Replace `input_file.xls` with the name of the Excel file you want to convert.

## xlsToCsvTerminalSCP.py

This script converts Excel files to CSV and automatically transfers them via SCP to a remote server.

### Requirements

This script only works on Linux or WSL (for Windows users). You will need the following dependencies:

### Python Packages:
- `pandas`
- `python-magic`
- `tqdm`

### Install dependencies:

```bash
pip install pandas python-magic tqdm
```

### System Requirement:

You must have `sshpass` installed to transfer files via SCP.

- **On Ubuntu/Debian:**

```bash
sudo apt-get install sshpass
```

- **On macOS:**

```bash
brew install sshpass
```

To run the script:

```bash
python xlsToCsvTerminalSCP.py input_files*.xls
```

The script will prompt you to enter your SSH credentials to send the converted files. By default, files are transferred to the `/logs` directory on the remote server, but you can modify this by changing line 43 of the code.

**Note:** Both the sender and receiver machines must have `sshpass` installed.
