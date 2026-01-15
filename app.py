import pandas as pd
import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

# === Setup GUI for file/folder selection ===
root = tk.Tk()
root.withdraw()  # Hide main window

# Select Excel file
excel_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx *.xls")])
if not excel_path:
    messagebox.showerror("Error", "No Excel file selected.")
    exit()

# Select images folder
images_folder = filedialog.askdirectory(title="Select Folder Containing Images")
if not images_folder:
    messagebox.showerror("Error", "No image folder selected.")
    exit()

# Select destination folder
destination_folder = filedialog.askdirectory(title="Select Destination Folder")
if not destination_folder:
    messagebox.showerror("Error", "No destination folder selected.")
    exit()

# === Load image names from Excel ===
try:
    df = pd.read_excel(excel_path)
    image_names = df.iloc[:, 0].astype(str).tolist()
except Exception as e:
    messagebox.showerror("Error", f"Failed to read Excel file: {e}")
    exit()

# === Copy Images ===
missing_images = []

for img_name in image_names:
    src = os.path.join(images_folder, img_name)
    dst = os.path.join(destination_folder, img_name)
    if os.path.isfile(src):
        shutil.copy2(src, dst)
    else:
        missing_images.append(img_name)

# === Results ===
if missing_images:
    print("\n❌ The following images were NOT found in the folder:")
    for name in missing_images:
        print(f" - {name}")
else:
    print("\n✅ All images were found and copied successfully.")

print("\n✅ Process complete.")
