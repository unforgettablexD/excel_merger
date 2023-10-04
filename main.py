import tkinter as tk
from tkinter import filedialog
import pandas as pd

# Create the main Tkinter window
root = tk.Tk()
root.title("Excel File Joiner")

# Function to select Excel files
def select_files():
    files = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx")])
    if files:
        process_files(files)

# Function to join Excel files
def process_files(files):
    common_columns = set()
    dataframes = []
    
    for file_path in files:
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            tk.messagebox.showerror("Error", f"Error reading {file_path}: {str(e)}")
            continue
        
        if common_columns:
            if set(df.columns) != common_columns:
                tk.messagebox.showerror("Error", f"Column names in {file_path} do not match the previous files.")
                return
        else:
            common_columns = set(df.columns)
        
        dataframes.append(df)

    if not dataframes:
        tk.messagebox.showinfo("Info", "No valid Excel files selected.")
        return

    joined_df = pd.concat(dataframes, ignore_index=True)
    
    # Create a new window to display the joined DataFrame
    result_window = tk.Toplevel(root)
    result_window.title("Joined DataFrame")
    text = tk.Text(result_window)
    text.insert("1.0", joined_df.to_string(index=False))
    text.pack()
    
    # Save the joined DataFrame to a new Excel file
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if save_path:
        try:
            joined_df.to_excel(save_path, index=False)
            tk.messagebox.showinfo("Info", f"Joined data saved to {save_path}")
        except Exception as e:
            tk.messagebox.showerror("Error", f"Error saving the joined data: {str(e)}")

# Create UI elements
select_button = tk.Button(root, text="Select Excel Files", command=select_files)
select_button.pack()

# Start the main loop
root.mainloop()