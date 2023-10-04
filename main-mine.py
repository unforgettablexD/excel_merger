import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd

# Create the main Tkinter window
root = tk.Tk()
root.title("File Joiner")

selected_files = []
column_dropdowns = []

# Function to select files
def select_files():
    global selected_files, column_dropdowns
    file_types = [("CSV Files", "*.csv"), ("Excel Files", "*.xlsx"), ("JSON Files", "*.json")]
    selected_files = filedialog.askopenfilenames(filetypes=file_types)
    if selected_files:
        # Clear previous dropdowns
        for dropdown in column_dropdowns:
            dropdown.destroy()
        column_dropdowns.clear()
        
        # Create new dropdowns for each file
        for file_path in selected_files:
            create_column_dropdown(file_path)

# Create a dropdown for the columns of a file
def create_column_dropdown(file_path):
    global column_dropdowns
    try:
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.json'):
            df = pd.read_json(file_path)
        else:
            return
        column_names = df.columns.tolist()
        column_var = tk.StringVar()
        column_var.set(column_names[0])
        column_dropdown = ttk.Combobox(root, textvariable=column_var, values=column_names)
        column_dropdown.pack(pady=5)
        column_dropdowns.append(column_dropdown)
    except Exception as e:
        messagebox.showerror("Error", f"Error reading {file_path}: {str(e)}")

# Function to handle overlapping columns
def handle_overlapping_columns(joined_df):
    overlapping_columns = [col for col in joined_df.columns if col.endswith('_x') or col.endswith('_y')]
    overlapping_base = set([col[:-2] for col in overlapping_columns])
    
    if not overlapping_base:
        return joined_df

    def combine_columns():
        for col in overlapping_base:
            col_x = col + '_x'
            col_y = col + '_y'
            if var_dict[col].get():
                joined_df[col] = joined_df[col_x].where(joined_df[col_x].notnull(), joined_df[col_y])
            joined_df.drop([col_x, col_y], axis=1, inplace=True)
        combine_window.destroy()
        display_data(joined_df)

    combine_window = tk.Toplevel(root)
    combine_window.title("Combine Columns")

    var_dict = {}
    for col in overlapping_base:
        var = tk.BooleanVar()
        chk = tk.Checkbutton(combine_window, text=f"Combine {col} columns?", variable=var)
        chk.pack(pady=5)
        var_dict[col] = var

    combine_button = tk.Button(combine_window, text="Combine Selected Columns", command=combine_columns)
    combine_button.pack(pady=10)

# Function to display the joined data in a Treeview
def display_data(df):
    preview_window = tk.Toplevel(root)
    preview_window.title("Preview Joined Data")
    
    tree = ttk.Treeview(preview_window, columns=list(df.columns), show='headings')
    tree.pack(fill=tk.BOTH, expand=True)
    
    for column in df.columns:
        tree.heading(column, text=column)
        tree.column(column, width=100)
    
    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))

    save_button = tk.Button(preview_window, text="Save Data", command=lambda: save_data(df))
    save_button.pack(pady=10)

# Function to save the joined data
def save_data(df):
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if save_path:
        try:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Info", f"Joined data saved to {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving the joined data: {str(e)}")

# Function to join files and display the preview
def process_files():
    global selected_files, column_dropdowns
    dataframes = []
    join_columns = [dropdown.get() for dropdown in column_dropdowns]
    
    for file_path in selected_files:
        try:
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            elif file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            elif file_path.endswith('.json'):
                df = pd.read_json(file_path)
            else:
                continue
        except Exception as e:
            messagebox.showerror("Error", f"Error reading {file_path}: {str(e)}")
            continue
        
        dataframes.append(df)

    if not dataframes:
        messagebox.showinfo("Info", "No valid files selected.")
        return

    join_type = join_type_var.get()
    joined_df = dataframes[0]
    for i, df in enumerate(dataframes[1:], 1):
        joined_df = pd.merge(joined_df, df, how=join_type, left_on=join_columns[0], right_on=join_columns[i])

    handle_overlapping_columns(joined_df)

# UI elements
instructions = tk.Label(root, text="Select files to join and choose the join type.")
instructions.pack(pady=10)

join_type_var = tk.StringVar()
join_type_var.set("inner")
join_type_label = tk.Label(root, text="Join Type:")
join_type_label.pack(pady=5)
join_type_dropdown = ttk.Combobox(root, textvariable=join_type_var, values=["inner", "outer", "left"])
join_type_dropdown.pack(pady=5)

select_button = tk.Button(root, text="Select Files", command=select_files)
select_button.pack(pady=10)

process_button = tk.Button(root, text="Process Files", command=process_files)
process_button.pack(pady=10)

# Start the main loop
root.mainloop()
