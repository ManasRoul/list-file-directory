# File Listing Script

Export a complete list of files from any local directory to Excel with metadata.

## How It Works

**The script scans from wherever you run it:**
- Run in a directory → Lists all files in that directory and subdirectories
- Run at drive root → Lists all files in the entire drive
- Specify path with `-d` → Lists all files from that location

## Installation

```bash
# Create and activate virtual environment
python3 -m venv venv
source venv/bin/activate

# Install dependencies
pip install pandas openpyxl
```

## Quick Start

```bash
# Always activate virtual environment first
source venv/bin/activate

# Then run the script:

# Scan CURRENT directory (wherever you are)
python script/list.py

# Scan from a SPECIFIC directory
python script/list.py -d /path/to/directory

# Scan entire DRIVE
python script/list.py -d /Volumes/MyDrive -o drive_files.xlsx
```

### Scan Local Disk Locations

```bash
# OPTION 1: Navigate to the directory first, then run
cd ~/Downloads
python /path/to/script/list.py -o downloads_files.xlsx

# OPTION 2: Specify the directory with -d flag
python script/list.py -d ~/Downloads -o downloads_files.xlsx

# Scan your Documents folder
python script/list.py -d ~/Documents -o documents_files.xlsx

# Scan your Desktop
python script/list.py -d ~/Desktop -o desktop_files.xlsx

# Scan entire home directory (⚠️ may take a while)
python script/list.py -d ~ -o home_all_files.xlsx

# Scan a specific project folder
python script/list.py -d /Users/username/Projects/MyProject -o project_files.xlsx

# Scan external drive or USB (entire drive)
python script/list.py -d /Volumes/MyUSB -o usb_files.xlsx

# Scan network drive
python script/list.py -d /Volumes/NetworkDrive -o network_files.xlsx

# Scan with custom exclusions
python script/list.py -d ~/Documents -e .git node_modules temp cache -o docs_files.xlsx
```

### Real-World Scenarios

```bash
# Clean up Downloads folder - see what's there
cd ~/Downloads
python /path/to/checking/script/list.py -o downloads_analysis.xlsx

# Audit entire system drive (⚠️ will take hours)
python script/list.py -d / -o system_drive_audit.xlsx

# Find all files on external backup drive
python script/list.py -d /Volumes/Backup2024 -o backup_inventory.xlsx

# Analyze project without certain folders
python script/list.py -d ~/Projects/BigApp -e node_modules dist build cache .git -o project_clean.xlsx
```

## Options

- `-d, --directory`: Directory to scan (default: current directory)
- `-o, --output`: Output Excel file path (default: file_list.xlsx)
- `-e, --exclude`: Directories to exclude from scanning

## Output

The script creates an Excel file with 3 sheets:

1. **All Files**: Complete listing of all files with:
   - File name
   - Extension
   - Directory path
   - Relative path
   - Full path
   - File size (bytes and human-readable)
   - Modified date
   - Created date

2. **Summary**: Overall statistics including:
   - Total file count
   - Total size
   - Number of unique extensions
   - Report generation timestamp

3. **Extension Summary**: Files grouped by extension with counts and sizes

## Examples for Different Scenarios

### Find all PDFs in your system
After generating the Excel file, use Excel's filter feature on the "Extension" column to show only `.pdf` files.

### Analyze disk usage by file type
Check the "Extension Summary" sheet to see which file types are taking up the most space.

### Backup documentation
Generate a file list before backing up to track what files existed at that point in time.
