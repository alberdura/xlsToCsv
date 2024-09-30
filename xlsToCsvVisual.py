import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import magic

class FileConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("XLS to CSV")
        self.geometry("500x300")
        self.configure(bg="#ab79fb")

        self.frame = tk.Frame(self, bg="#c7a6fc", bd=5, relief="raised")
        self.frame.place(relx=0.5, rely=0.5, anchor="center", width=450, height=200)

        self.label = tk.Label(self.frame, text="Select an .xls, .xlsx or .csv file", font=("Arial", 12, "bold"), bg="#ffffff")
        self.label.pack(pady=1)
        self.label = tk.Label(self.frame, text="(The output file will be saved in the same location as the input file)", font=("Arial", 8, "bold"), bg="#ffffff")
        self.label.pack(pady=15)

        self.select_button = tk.Button(self.frame, text="Select File", command=self.select_file, font=("Arial", 10, "bold"), bg="#4c4184", fg="#ffffff")
        self.select_button.pack(pady=10)

        self.progress = ttk.Progressbar(self.frame, orient="horizontal", length=300, mode="indeterminate")
        self.progress.pack(pady=20)

        self.close_button = tk.Button(self.frame, text="Close", command=self.quit, font=("Arial", 10, "bold"), bg="#007bff", fg="#ffffff")
        self.close_button.pack(pady=10)

        self.selected_file_path = None

    def select_file(self):
        self.selected_file_path = filedialog.askopenfilename(
            title="Select File",
            filetypes=[("Excel Files", "*.xls *.xlsx"), ("CSV Files", "*.csv")]
        )
        
        if self.selected_file_path:
            self.progress.start()
            threading.Thread(target=self.process_file, args=(self.selected_file_path,), daemon=True).start()

    def process_file(self, file_path):
        file_type = magic.from_file(file_path, mime=True)

        try:
            if file_type in ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']:
                csv_file_path = os.path.splitext(file_path)[0] + '.csv'
                df = pd.read_excel(file_path)
                df.to_csv(csv_file_path, index=False)
                self.show_message("Success", f"Excel file converted to {csv_file_path}")
            elif file_type == 'text/plain':
                df = pd.read_csv(file_path, on_bad_lines='skip')
                csv_file_path = os.path.splitext(file_path)[0] + '_converted.csv'
                df.to_csv(csv_file_path, index=False)
                self.show_message("Success", f"CSV file converted to {csv_file_path}")
            else:
                self.show_message("Warning", "The file is neither a valid Excel nor CSV file. Detected type: " + file_type)
        except Exception as e:
            self.show_message("Error", f"Error converting file: {e}")
        finally:
            self.progress.stop()

    def show_message(self, title, message):
        self.after(0, lambda: messagebox.showinfo(title, message) if title == "Success" else messagebox.showwarning(title, message) if title == "Warning" else messagebox.showerror(title, message))

if __name__ == "__main__":
    app = FileConverterApp()
    app.mainloop()
