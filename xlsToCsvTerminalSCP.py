import os
import pandas as pd
import magic
import argparse
import glob
from tqdm import tqdm
import subprocess

def process_file(file_path):
    file_type = magic.from_file(file_path, mime=True)

    try:
        if file_type in ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
            csv_file_path = os.path.splitext(file_path)[0] + '_converted.csv'
            total_rows = sum([chunk.shape[0] for chunk in pd.read_excel(file_path, chunksize=10000)])
            
            with open(csv_file_path, 'w', encoding='utf-8') as f:
                with tqdm(total=total_rows, desc=f"Converting {os.path.basename(file_path)}", unit="row") as pbar:
                    for chunk in pd.read_excel(file_path, chunksize=10000):
                        chunk.to_csv(f, header=f.tell() == 0, index=False, encoding='utf-8')
        
        elif file_type == 'text/plain':
            csv_file_path = os.path.splitext(file_path)[0] + '_converted.csv'
            total_rows = sum([chunk.shape[0] for chunk in pd.read_csv(file_path, chunksize=10000, on_bad_lines='skip')])

            with open(csv_file_path, 'w', encoding='utf-8') as f:
                with tqdm(total=total_rows, desc=f"Converting {os.path.basename(file_path)}", unit="row") as pbar:
                    for chunk in pd.read_csv(file_path, chunksize=10000, on_bad_lines='skip'):
                        chunk.to_csv(f, header=f.tell() == 0, index=False, encoding='utf-8')
        else:
            print(f"\nWarning: The file {file_path} is neither a valid Excel nor CSV. Detected type: {file_type}")
    
    except Exception as e:
        print(f"\nError converting the file {file_path}: {e}")

def scp_transfer(files, username, ip, pssw):
    prefix = os.path.basename(files[0])[:4]
    files_to_transfer = [os.path.join(os.path.dirname(os.path.abspath(prefix)), f"{os.path.basename(prefix)[:4]}*.csv")]
    scp_command = f"sshpass -p '{pssw}' scp {' '.join(files_to_transfer)} {username}@{ip}:/home/{username}/logs"

    try:
        subprocess.run(scp_command, shell=True, check=True)
        print(f"Files successfully transferred to {username}@{ip}")
    except subprocess.CalledProcessError as e:
        print(f"Error transferring the files: {e}")

def main():
    username = input("User: ")
    ip = input("IP: ")
    pssw = input("Pssw: ")

    parser = argparse.ArgumentParser(description='Convert XLS or XLSX files to CSV.')
    parser.add_argument('files', metavar='F', type=str, nargs='+', help='File patterns to convert')
    
    args = parser.parse_args()
    converted_files = []

    for pattern in args.files:
        file_paths = glob.glob(pattern)
        if not file_paths:
            print(f"No files found matching the pattern: {pattern}")
        else:
            for file_path in file_paths:
                if os.path.exists(file_path):
                    process_file(file_path)
                    converted_files.append(file_path)
                else:
                    print(f"Error: The file {file_path} does not exist.")
    
    if converted_files:
        scp_transfer(converted_files, username, ip, pssw)

if __name__ == "__main__":
    main()
