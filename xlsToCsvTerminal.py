import os
import pandas as pd
import magic
import argparse
import glob

def process_file(file_path):
    file_type = magic.from_file(file_path, mime=True)

    try:
        if file_type in ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
            df = pd.read_excel(file_path)
            csv_file_path = os.path.splitext(file_path)[0] + '.csv'
            df.to_csv(csv_file_path, index=False, encoding='utf-8')
            print(f"Success: Excel file converted to {csv_file_path}")
        elif file_type == 'text/plain':
            df = pd.read_csv(file_path, on_bad_lines='skip')
            csv_file_path = os.path.splitext(file_path)[0] + '_converted.csv'
            df.to_csv(csv_file_path, index=False, encoding='utf-8')
            print(f"Success: CSV file converted to {csv_file_path}")
        else:
            print(f"Warning: The file {file_path} is neither a valid Excel nor CSV file. Detected type: {file_type}")
    except Exception as e:
        print(f"Error converting the file {file_path}: {e}")

def main():
    parser = argparse.ArgumentParser(description='Convert XLS or XLSX files to CSV.')
    parser.add_argument('files', metavar='F', type=str, nargs='+', help='File patterns to convert')
    
    args = parser.parse_args()

    for pattern in args.files:
        file_paths = glob.glob(pattern)
        if not file_paths:
            print(f"No files found matching the pattern: {pattern}")
        else:
            for file_path in file_paths:
                if os.path.exists(file_path):
                    process_file(file_path)
                else:
                    print(f"Error: The file {file_path} does not exist.")

if __name__ == "__main__":
    main()