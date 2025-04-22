import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import datetime
import re
from tkinter import PhotoImage

# Get desktop path and create SORT RESULT folder
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
sort_result_folder = os.path.join(desktop_path, "SORT RESULT")

if not os.path.exists(sort_result_folder):
    os.makedirs(sort_result_folder)

# Create the main application window
root = tk.Tk()
root.title("Mkononi Numbers-Sorting System")
logo = PhotoImage(file=r"C:\Goddie\Numbers-sort\icon.png")
root.tk.call('wm', 'iconphoto', root._w, logo)
root.resizable(False, False)

# Define StringVar for folder path and search name
folder_var = tk.StringVar()
search_name_var = tk.StringVar()

# Create UI elements
tk.Label(root, text="Select Folder:").pack(pady=5)
tk.Entry(root, textvariable=folder_var, width=50).pack(pady=5)
tk.Button(root, text="Browse", command=lambda: folder_var.set(filedialog.askdirectory())).pack(pady=5)

tk.Label(root, text="Search Name(s):").pack(pady=5)
tk.Entry(root, textvariable=search_name_var).pack(pady=5)

# Global progress bar (created once)
progress = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=400)
progress.pack(pady=10, fill='x', padx=20)

tk.Button(root, text="Execute Filter", command=lambda: execute_filter()).pack(pady=20)

# Progress update function
def update_progress(value):
    progress['value'] = value
    root.update_idletasks()

def execute_filter():
    folder = folder_var.get()
    search_input = search_name_var.get().strip().lower()

    search_names = list(set(name.strip() for name in re.split(r'[,\n\r]+', search_input) if name.strip()))

    if not search_names:
        messagebox.showwarning("Input Error", "Please enter at least one name to search.")
        return

    required_columns = ['Credit Identity String', 'Customer Name']
    files_processed = 0

    # Get only supported files
    supported_files = [f for f in os.listdir(folder) if f.endswith(('.xlsx', '.xlsm', '.csv'))]
    progress['value'] = 0
    progress['maximum'] = len(supported_files)

    # Dictionary to store collected data for each search name
    name_data = {name: pd.DataFrame(columns=required_columns) for name in search_names}

    try:
        for filename in supported_files:
            filepath = os.path.join(folder, filename)
            files_processed += 1
            update_progress(files_processed)

            if filename.endswith(('.xlsx', '.xlsm')):
                xls = pd.ExcelFile(filepath)
                for sheet_name in xls.sheet_names:
                    df = xls.parse(sheet_name)
                    df.columns = df.columns.str.strip()
                    actual_columns = df.columns.str.lower()

                    if all(col.lower() in actual_columns for col in required_columns):
                        for search_name in search_names:
                            pattern = rf'\b{re.escape(search_name)}\b'
                            rows = df[df['Customer Name'].str.contains(pattern, case=False, na=False, regex=True)]
                            if not rows.empty:
                                available_cols = [col for col in required_columns if col in df.columns]
                                name_data[search_name] = pd.concat([name_data[search_name], rows[available_cols]], ignore_index=True)

            elif filename.endswith('.csv'):
                try:
                    df = pd.read_csv(filepath, on_bad_lines='skip')
                    df.columns = df.columns.str.strip()
                    actual_columns = df.columns.str.lower()

                    if all(col.lower() in actual_columns for col in required_columns):
                        for search_name in search_names:
                            pattern = rf'\b{re.escape(search_name)}\b'
                            rows = df[df['Customer Name'].str.contains(pattern, case=False, na=False, regex=True)]
                            if not rows.empty:
                                available_cols = [col for col in required_columns if col in df.columns]
                                name_data[search_name] = pd.concat([name_data[search_name], rows[available_cols]], ignore_index=True)

                except pd.errors.ParserError as e:
                    print(f"Error parsing CSV: {filename} - {e}")

        matches_found = 0
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        for name, df in name_data.items():
            if not df.empty:
                output_file = os.path.join(sort_result_folder, f"{name}_{timestamp}.xlsx")
                df.to_excel(output_file, index=False)
                print(f"Saved: {output_file}")
                matches_found += 1

        if matches_found > 0:
            messagebox.showinfo("Done", f"Created {matches_found} result files.")
        else:
            messagebox.showwarning("No Matches", "No matches found for the given names.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

    finally:
        update_progress(0)

# Run the application
root.mainloop()
