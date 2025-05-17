import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import sys
import os

def select_source():
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if filepath:
        source_entry.delete(0, tk.END)
        source_entry.insert(0, filepath)

def run_scripts():
    source = source_entry.get()

    if not os.path.exists(source):
        messagebox.showerror("Error", "Source file does not exist.")
        return

    # Extract the base name and directory from the source file
    source_dir = os.path.dirname(source)
    base_name = os.path.splitext(os.path.basename(source))[0]

    # Define the output file names using the same directory and base name with suffixes
    output_capacity = os.path.join(source_dir, f"{base_name}_capacity.xlsx")
    output_individual = os.path.join(source_dir, f"{base_name}_individual.xlsx")
    output_skillmatrix = os.path.join(source_dir, f"{base_name}_skillmatrix.xlsx")

    try:
        # Run the three scripts, passing source and output as arguments
        subprocess.run([sys.executable, "build-capacity.py", source, output_capacity], check=True)
        subprocess.run([sys.executable, "build-individual.py", source, output_individual], check=True)
        subprocess.run([sys.executable, "build-skillmatrix.py", source, output_skillmatrix], check=True)

        messagebox.showinfo("Success", "All scripts executed successfully!")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Execution Failed", str(e))

# GUI setup
root = tk.Tk()
root.title("Skill Matrix Generator by Nino Yuliantoro")

tk.Label(root, text="Source File:").grid(row=0, column=0, sticky="e")
source_entry = tk.Entry(root, width=50)
source_entry.grid(row=0, column=1)
tk.Button(root, text="Browse", command=select_source).grid(row=0, column=2)

tk.Button(root, text="OK", command=run_scripts, width=20).grid(row=1, column=0, columnspan=3, pady=10)

root.mainloop()
