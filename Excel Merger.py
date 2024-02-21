import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd

def merge_excel_files(file1, file2, output_file):
    try:
        # Read Excel files into pandas DataFrames
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # Merge DataFrames
        merged_df = pd.concat([df1, df2], ignore_index=True)

        # Write merged DataFrame to a new Excel file
        merged_df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"Merged files '{file1}' and '{file2}' successfully into '{output_file}'")
    except Exception as e:
        print(f"Error merging files: {e}")

def browse_excel_file(entry_widget):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

def merge_files():
    file1 = entry_file1.get()
    file2 = entry_file2.get()
    output_file = entry_output_file.get()

    if file1 and file2 and output_file:  # Check if all file paths are non-empty
        merge_excel_files(file1, file2, output_file)
    else:
        print("Please select all files and specify an output file.")

root = tk.Tk()
root.title("Excel File Merger")

# Dark mode color scheme
bg_color = "#1e1e1e"
fg_color = "#ffffff"
font = ("Arial", 10)

root.config(bg=bg_color)

# File 1
label_file1 = tk.Label(root, text="Select File 1:", font=font, bg=bg_color, fg=fg_color)
label_file1.grid(row=0, column=0, padx=10, pady=5)
entry_file1 = tk.Entry(root, width=40, font=font, bg=bg_color, fg=fg_color)
entry_file1.grid(row=0, column=1, padx=10, pady=5)
button_browse1 = ttk.Button(root, text="Browse", command=lambda: browse_excel_file(entry_file1))
button_browse1.grid(row=0, column=2, padx=10, pady=5)

# File 2
label_file2 = tk.Label(root, text="Select File 2:", font=font, bg=bg_color, fg=fg_color)
label_file2.grid(row=1, column=0, padx=10, pady=5)
entry_file2 = tk.Entry(root, width=40, font=font, bg=bg_color, fg=fg_color)
entry_file2.grid(row=1, column=1, padx=10, pady=5)
button_browse2 = ttk.Button(root, text="Browse", command=lambda: browse_excel_file(entry_file2))
button_browse2.grid(row=1, column=2, padx=10, pady=5)

# Output File
label_output_file = tk.Label(root, text="Output File:", font=font, bg=bg_color, fg=fg_color)
label_output_file.grid(row=2, column=0, padx=10, pady=5)
entry_output_file = tk.Entry(root, width=40, font=font, bg=bg_color, fg=fg_color)
entry_output_file.grid(row=2, column=1, padx=10, pady=5)

# Merge Button
button_merge = ttk.Button(root, text="Merge Files", command=merge_files)
button_merge.grid(row=3, column=1, padx=10, pady=10)

root.mainloop()
