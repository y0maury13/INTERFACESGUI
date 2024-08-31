import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

# Body App
root = tk.Tk()
root.geometry("800x600")
root.pack_propagate(False)
root.resizable(True, True)
root.title("Excel Data Viewer © All copyrights reserved by Yomaury 2024")  # Title with copyright

# Function to toggle fullscreen mode
def toggle_fullscreen(event=None):
    state = not root.attributes('-fullscreen')
    root.attributes('-fullscreen', state)
    if state:
        root.geometry("")  # Remove fixed geometry to allow full expansion
    else:
        root.geometry("800x600")  # Revert to a specific size

# Bind F11 key to toggle fullscreen
root.bind("<F11>", toggle_fullscreen)

# Set background color
root.configure(bg="#808080")

# Frame for Treeview
frame1 = tk.LabelFrame(root, text="Excel Data", bg="#e6e6e6", fg="black", font=("Arial", 12, "italic"))
frame1.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.6)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File", bg="#e6e6e6", fg="black", font=("Arial", 12, "italic"))
file_frame.place(relx=0.02, rely=0.65, relwidth=0.96, relheight=0.2)

# Button style
button_style = {
    "bg": "#2c2c2c",  # Green background
    "fg": "white",    # White text
    "font": ("Arial", 10, "italic"),
    "relief": tk.RAISED
}

# Buttons
button1 = tk.Button(file_frame, text="Browse a file", command=lambda: file_dialog(), **button_style)
button1.place(rely=0.65, relx=0.50)

button2 = tk.Button(file_frame, text="Load File", command=lambda: load_exceldata(), **button_style)
button2.place(rely=0.65, relx=0.30)

button3 = tk.Button(file_frame, text="Edit File", command=lambda: edit_exceldata(), **button_style)
button3.place(rely=0.65, relx=0.10)

label_file = ttk.Label(file_frame, text="No File Selected", background="#e6e6e6", font=("Arial", 10))
label_file.place(rely=0, relx=0)

# TREEVIEW WIDGET
tv1 = ttk.Treeview(frame1, show="tree headings")
tv1.place(relheight=1, relwidth=1)

# Scrollbars
treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview)
treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview)
tv1.config(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
treescrollx.pack(side="bottom", fill="x")
treescrolly.pack(side="right", fill="y")

# Global variable to store the DataFrame
df = None

# Functions
def file_dialog():
    filename = filedialog.askopenfilename(
        initialdir="/",
        title="Select a File",
        filetypes=(("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*"))
    )
    label_file["text"] = filename

def load_exceldata():
    global df
    file_path = label_file["text"]

    try:
        excel_filename = file_path
        if excel_filename.endswith(".xlsx"):
            df = pd.read_excel(excel_filename)
        elif excel_filename.endswith(".csv"):
            df = pd.read_csv(excel_filename)
        else:
            messagebox.showerror("Information", "The selected file is not supported")
            return None
        clear_data()

        # Set column headings and insert data
        tv1["columns"] = list(df.columns)
        tv1["show"] = "headings"
        for column in tv1["columns"]:
            tv1.heading(column, text=column)

        df_rows = df.to_numpy().tolist()
        for i, row in enumerate(df_rows):
            tv1.insert("", "end", text=f"Row {i+1}", values=row)

    except ValueError:
        messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        messagebox.showerror("Information", f"No such file as {file_path}")
        return None

def edit_exceldata():
    global df
    if df is None:
        messagebox.showwarning("Warning", "You need to load a file first!")
        return

    selected_item = tv1.selection()
    if not selected_item:
        messagebox.showwarning("Warning", "Please select a row to edit")
        return

    row = tv1.item(selected_item, 'values')
    col_names = list(df.columns)
    
    edit_window = tk.Toplevel(root)
    edit_window.title("Edit Row")
    
    entry_vars = []
    for i, value in enumerate(row):
        label = tk.Label(edit_window, text=f"{col_names[i]}:")
        label.grid(row=i, column=0, padx=10, pady=5)
        
        entry_var = tk.StringVar(value=value)
        entry = tk.Entry(edit_window, textvariable=entry_var)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entry_vars.append(entry_var)

    def save_changes():
        new_values = [entry_var.get() for entry_var in entry_vars]
        for i, col in enumerate(col_names):
            df.at[int(selected_item[0]), col] = new_values[i]
        
        load_exceldata()  # Refresh the Treeview
        save_to_file()  # Save changes to the file
        edit_window.destroy()

    save_button = tk.Button(edit_window, text="Save", command=save_changes, **button_style)
    save_button.grid(row=len(row), columnspan=2, pady=10)

def save_to_file():
    global df
    file_path = label_file["text"]
    try:
        if file_path.endswith(".xlsx"):
            df.to_excel(file_path, index=False)
        elif file_path.endswith(".csv"):
            df.to_csv(file_path, index=False)
        else:
            messagebox.showerror("Error", "Unsupported file format")
    except Exception as e:
        messagebox.showerror("Error", f"Could not save file: {str(e)}")

def clear_data():
    tv1.delete(*tv1.get_children())

root.mainloop()
