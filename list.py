#!/usr/bin/env python3
"""
Script to list all files with their extensions and export to Excel.
Includes file metadata like size, modification date, and full path.
"""

import os
import pandas as pd
from pathlib import Path
from datetime import datetime
import argparse
import hashlib
from collections import defaultdict


def get_file_size_readable(size_bytes):
    """Convert bytes to human-readable format."""
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} PB"


def get_file_hash(file_path, block_size=65536):
    """Calculate MD5 hash of a file for exact duplicate detection."""
    try:
        hasher = hashlib.md5()
        with open(file_path, 'rb') as f:
            while True:
                block = f.read(block_size)
                if not block:
                    break
                hasher.update(block)
        return hasher.hexdigest()
    except (PermissionError, OSError, IOError):
        return None


def list_files_with_extensions(root_dir, exclude_dirs=None, calculate_hashes=True):
    """
    Traverse directory and collect file information.
    
    Args:
        root_dir: Root directory to start scanning
        exclude_dirs: List of directory names to exclude (e.g., ['node_modules', '.git'])
        calculate_hashes: Whether to calculate file hashes for exact duplicate detection
    
    Returns:
        List of dictionaries containing file information
    """
    if exclude_dirs is None:
        exclude_dirs = ['.git', '__pycache__', 'node_modules', '.venv', 'venv']
    
    file_data = []
    root_path = Path(root_dir).resolve()
    
    # Check if scanning from root/drive
    is_drive_root = len(root_path.parts) <= 2 or str(root_path).startswith('/Volumes/')
    
    print(f"\n📁 Scanning: {root_path}")
    if is_drive_root:
        print("   ⚠️  This appears to be a drive/volume root - will scan entire drive!")
    print(f"   Excluding: {', '.join(exclude_dirs)}")
    print("-" * 70)
    
    file_count = 0
    
    for dirpath, dirnames, filenames in os.walk(root_path):
        # Exclude specified directories
        dirnames[:] = [d for d in dirnames if d not in exclude_dirs]
        
        for filename in filenames:
            try:
                file_path = Path(dirpath) / filename
                
                # Get file stats
                stats = file_path.stat()
                
                # Extract extension
                extension = file_path.suffix if file_path.suffix else '(no extension)'
                
                # Calculate relative path
                try:
                    relative_path = file_path.relative_to(root_path)
                except ValueError:
                    relative_path = file_path
                
                # Calculate file hash for exact duplicate detection
                file_hash = None
                if calculate_hashes and stats.st_size > 0:  # Skip empty files
                    file_hash = get_file_hash(file_path)
                
                file_info = {
                    'File Name': filename,
                    'Extension': extension,
                    'Directory': str(file_path.parent),
                    'Relative Path': str(relative_path),
                    'Full Path': str(file_path),
                    'Size (Bytes)': stats.st_size,
                    'Size (Readable)': get_file_size_readable(stats.st_size),
                    'Modified Date': datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                    'Created Date': datetime.fromtimestamp(stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                    'File Hash': file_hash,
                }
                
                file_data.append(file_info)
                file_count += 1
                
                # Show progress every 100 files
                if file_count % 100 == 0:
                    print(f"   Processed {file_count} files...", end='\r')
                
            except (PermissionError, OSError) as e:
                # Silently skip permission errors in large scans
                continue
    
    print(f"   Processed {file_count} files - Complete!     ")
    
    return file_data


def export_to_excel(file_data, output_file):
    """
    Export file data to Excel with formatting.
    
    Args:
        file_data: List of dictionaries containing file information
        output_file: Path to output Excel file
    """
    if not file_data:
        print("No files found to export!")
        return
    
    # Create DataFrame
    df = pd.DataFrame(file_data)
    
    # Detect duplicates by filename
    filename_counts = df['File Name'].value_counts()
    df['Duplicate by Name'] = df['File Name'].map(lambda x: 'Yes' if filename_counts[x] > 1 else 'No')
    df['Name Duplicate Count'] = df['File Name'].map(filename_counts)
    
    # Detect exact duplicates by hash (same content)
    if 'File Hash' in df.columns:
        hash_counts = df['File Hash'].value_counts()
        df['Exact Duplicate'] = df['File Hash'].map(lambda x: 'Yes' if pd.notna(x) and hash_counts[x] > 1 else 'No')
        df['Exact Duplicate Count'] = df['File Hash'].map(lambda x: hash_counts[x] if pd.notna(x) else 1)
        
        # Create duplicate groups for exact duplicates
        hash_to_group = {}
        group_id = 1
        for hash_val in df[df['Exact Duplicate'] == 'Yes']['File Hash'].unique():
            if pd.notna(hash_val):
                hash_to_group[hash_val] = f"DUP-{group_id}"
                group_id += 1
        df['Duplicate Group'] = df['File Hash'].map(lambda x: hash_to_group.get(x, ''))
    
    # Reorder columns to put duplicate info near the front
    cols = df.columns.tolist()
    # Move duplicate columns after File Name
    priority_cols = ['File Name', 'Duplicate by Name', 'Name Duplicate Count', 'Exact Duplicate', 'Exact Duplicate Count', 'Duplicate Group']
    other_cols = [c for c in cols if c not in priority_cols]
    df = df[[c for c in priority_cols if c in cols] + other_cols]
    
    # Sort by duplicate status, then extension, then file name
    df = df.sort_values(['Duplicate by Name', 'Extension', 'File Name'], ascending=[False, True, True])
    
    # Create summary statistics
    summary_data = {
        'Total Files': [len(df)],
        'Total Size (Bytes)': [df['Size (Bytes)'].sum()],
        'Total Size (Readable)': [get_file_size_readable(df['Size (Bytes)'].sum())],
        'Unique Extensions': [df['Extension'].nunique()],
        'Report Generated': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
    }
    summary_df = pd.DataFrame(summary_data)
    
    # Extension summary
    ext_summary = df.groupby('Extension').agg({
        'File Name': 'count',
        'Size (Bytes)': 'sum'
    }).reset_index()
    ext_summary.columns = ['Extension', 'File Count', 'Total Size (Bytes)']
    ext_summary['Total Size (Readable)'] = ext_summary['Total Size (Bytes)'].apply(get_file_size_readable)
    ext_summary = ext_summary.sort_values('File Count', ascending=False)
    
    # Duplicate summary - files with same name
    duplicates_by_name = df[df['Duplicate by Name'] == 'Yes'].copy()
    
    # Exact duplicate summary - files with same content
    exact_duplicates = df[df['Exact Duplicate'] == 'Yes'].copy() if 'Exact Duplicate' in df.columns else pd.DataFrame()
    
    # Export to Excel with multiple sheets
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Main file list
        df.to_excel(writer, sheet_name='All Files', index=False)
        
        # Summary sheet
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Extension summary
        ext_summary.to_excel(writer, sheet_name='Extension Summary', index=False)
        
        # Duplicate files by name
        if not duplicates_by_name.empty:
            duplicates_by_name.to_excel(writer, sheet_name='Duplicates by Name', index=False)
        
        # Exact duplicates (same content)
        if not exact_duplicates.empty:
            exact_duplicates.to_excel(writer, sheet_name='Exact Duplicates', index=False)
        
        # Auto-adjust column widths
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"\n✓ Excel file created successfully: {output_file}")
    print(f"  - Total files: {len(df)}")
    print(f"  - Total size: {get_file_size_readable(df['Size (Bytes)'].sum())}")
    print(f"  - Unique extensions: {df['Extension'].nunique()}")
    print(f"  - Duplicate filenames: {len(duplicates_by_name)}")
    if not exact_duplicates.empty:
        wasted_space = exact_duplicates[exact_duplicates['Exact Duplicate Count'] > 1].groupby('File Hash')['Size (Bytes)'].first().sum() * (exact_duplicates.groupby('File Hash').size() - 1).sum() / exact_duplicates.groupby('File Hash').size().sum()
        print(f"  - Exact duplicates: {len(exact_duplicates)}")
        print(f"  - Wasted space from duplicates: {get_file_size_readable(exact_duplicates['Size (Bytes)'].sum() - exact_duplicates.groupby('File Hash')['Size (Bytes)'].first().sum())}")


def main():
    """Main function to run the script."""
    parser = argparse.ArgumentParser(
        description='List all files with extensions and export to Excel',
        epilog='''
Examples:
  # Scan current directory
  python %(prog)s
  
  # Scan specific directory
  python %(prog)s -d /Users/username/Documents
  
  # Scan entire home directory
  python %(prog)s -d ~ -o home_files.xlsx
  
  # Scan specific drive/volume
  python %(prog)s -d /Volumes/MyDrive -o mydrive_files.xlsx
  
  # Scan Downloads folder
  python %(prog)s -d ~/Downloads -o downloads_files.xlsx
        ''',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        '-d', '--directory',
        default='.',
        help='Directory to scan - can be any local path (default: current directory). Examples: /Users/username/Documents, ~, ~/Downloads, /Volumes/MyDrive'
    )
    parser.add_argument(
        '-o', '--output',
        default='file_list.xlsx',
        help='Output Excel file path (default: file_list.xlsx in current directory)'
    )
    parser.add_argument(
        '-e', '--exclude',
        nargs='*',
        default=['.git', '__pycache__', 'node_modules', '.venv', 'venv', 'Library', 'Applications'],
        help='Directories to exclude (default: .git __pycache__ node_modules .venv venv Library Applications)'
    )
    parser.add_argument(
        '--no-hash',
        action='store_true',
        help='Skip hash calculation for faster scanning (won\'t detect exact duplicates)'
    )
    
    args = parser.parse_args()
    
    # Expand user path (~ to home directory) and get absolute path
    root_dir = os.path.abspath(os.path.expanduser(args.directory))
    
    if not os.path.exists(root_dir):
        print(f"Error: Directory '{root_dir}' does not exist!")
        return
    
    if not os.path.isdir(root_dir):
        print(f"Error: '{root_dir}' is not a directory!")
        return
    
    # Warning for large directory scans
    large_dirs = ['/Users', '/System', '/Library', '/Volumes']
    if any(root_dir.startswith(d) and root_dir.count('/') <= d.count('/') + 1 for d in large_dirs):
        print("\n⚠️  WARNING: You're scanning a large directory!")
        print(f"   Directory: {root_dir}")
        print("   This may take several minutes and generate a large Excel file.")
        response = input("   Do you want to continue? (y/n): ")
        if response.lower() not in ['y', 'yes']:
            print("Scan cancelled.")
            return
    
    # List files
    print("=" * 50)
    print("FILE LISTING SCRIPT")
    print("=" * 50)
    calculate_hashes = not args.no_hash
    if not calculate_hashes:
        print("\n⚠️  Hash calculation disabled - exact duplicate detection unavailable")
    file_data = list_files_with_extensions(root_dir, args.exclude, calculate_hashes)
    
    if not file_data:
        print("\nNo files found in the specified directory.")
        return
    
    # Export to Excel
    print("\nExporting to Excel...")
    output_path = os.path.abspath(args.output)
    export_to_excel(file_data, output_path)
    
    print("\n" + "=" * 50)
    print(f"Full path: {output_path}")
    print("=" * 50)


if __name__ == "__main__":
    main()
