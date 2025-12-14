# Created by Bibhuti (Facebook.com/bibhutithecoolboy) If you find it useful, please consider donating something.
# Contact me on Facebook to get the donation details.
# If you need any custom tools, feel free to contact me.

# RAW File Sorter - Sorts RAW files based on matching JPEG files
# Supports drag and drop for folder selection.

import tkinter as tk
from tkinter import filedialog, ttk, messagebox, Toplevel
from tkinterdnd2 import DND_FILES, TkinterDnD
import shutil
import os
import webbrowser
from PIL import Image, ImageTk
import sys
import pandas as pd
import csv

class FileSorter:

    def __init__(self, root):

        self.root = root

        self.root.title("File Sorter")

        self.root.configure(bg="#f0f2f5")

        self.root.resizable(True, True)



        # Set window icon (replace with actual path to your icon)

        try:

            self.root.iconbitmap("app_icon.ico")

        except:

            pass



        # Configure styles

        self.style = ttk.Style()

        self.style.configure("TButton", font=("Arial", 10), padding=5)

        self.style.map("TButton", background=[("active", "#e0e0e0")])

        self.style.configure("TLabel", font=("Arial", 10), background="#f0f2f5")



        # Create main container

        self.main_frame = tk.Frame(root, bg="#f0f2f5", padx=15, pady=15)

        self.main_frame.pack(fill="both", expand=True)



        # Header

        tk.Label(

            self.main_frame,

            text="File Sorter",

            font=("Arial", 14, "bold"),

            bg="#f0f2f5",

            fg="#333333"

        ).grid(row=0, column=0, columnspan=3, pady=(0, 10))



        # Excel/CSV File selection

        tk.Label(

            self.main_frame,

            text="Select Excel/CSV File or Drag and Drop:",

            font=("Arial", 9),

            bg="#f0f2f5"

        ).grid(row=1, column=0, sticky="w", pady=2)



        self.entry_excel = tk.Entry(self.main_frame, width=35, font=("Arial", 9))

        self.entry_excel.grid(row=2, column=0, columnspan=2, sticky="ew", padx=(0, 5))

        self.entry_excel.drop_target_register(DND_FILES)

        self.entry_excel.dnd_bind('<<Drop>>', self.drop_excel_file)



        ttk.Button(

            self.main_frame,

            text="Browse",

            command=self.select_excel_file

        ).grid(row=2, column=2, sticky="w")



        # Excel/CSV File drag-and-drop area

        self.excel_drop = tk.Label(

            self.main_frame,

            text="Drop Excel/CSV File Here",

            font=("Arial", 9, "italic"),

            bg="#e6f3ff",

            bd=2,

            relief="groove",

            height=1,

            width=35

        )

        self.excel_drop.grid(row=3, column=0, columnspan=2, pady=2, padx=(0, 5))

        self.excel_drop.drop_target_register(DND_FILES)

        self.excel_drop.dnd_bind('<<Drop>>', self.drop_excel_file)



        # Image Folder selection

        tk.Label(

            self.main_frame,

            text="Select Image Folder or Drag and Drop:",

            font=("Arial", 9),

            bg="#f0f2f5"

        ).grid(row=4, column=0, sticky="w", pady=2)



        self.entry_image = tk.Entry(self.main_frame, width=35, font=("Arial", 9))

        self.entry_image.grid(row=5, column=0, columnspan=2, sticky="ew", padx=(0, 5))

        self.entry_image.drop_target_register(DND_FILES)

        self.entry_image.dnd_bind('<<Drop>>', self.drop_image_folder)



        ttk.Button(

            self.main_frame,

            text="Browse",

            command=self.select_image_folder

        ).grid(row=5, column=2, sticky="w")



        # Image Folder drag-and-drop area

        self.image_drop = tk.Label(

            self.main_frame,

            text="Drop Image Folder Here",

            font=("Arial", 9, "italic"),

            bg="#e6f3ff",

            bd=2,

            relief="groove",

            height=1,

            width=35

        )

        self.image_drop.grid(row=6, column=0, columnspan=2, pady=2, padx=(0, 5))

        self.image_drop.drop_target_register(DND_FILES)

        self.image_drop.dnd_bind('<<Drop>>', self.drop_image_folder)



        # Output Folder selection

        tk.Label(

            self.main_frame,

            text="Select Output Folder or Drag and Drop:",

            font=("Arial", 9),

            bg="#f0f2f5"

        ).grid(row=7, column=0, sticky="w", pady=2)



        self.entry_output = tk.Entry(self.main_frame, width=35, font=("Arial", 9))

        self.entry_output.grid(row=8, column=0, columnspan=2, sticky="ew", padx=(0, 5))

        self.entry_output.drop_target_register(DND_FILES)

        self.entry_output.dnd_bind('<<Drop>>', self.drop_output_folder)



        ttk.Button(

            self.main_frame,

            text="Browse",

            command=self.select_output_folder

        ).grid(row=8, column=2, sticky="w")



        # Output Folder drag-and-drop area

        self.output_drop = tk.Label(

            self.main_frame,

            text="Drop Output Folder Here",

            font=("Arial", 9, "italic"),

            bg="#e6f3ff",

            bd=2,

            relief="groove",

            height=1,

            width=35

        )

        self.output_drop.grid(row=9, column=0, columnspan=2, pady=2, padx=(0, 5))

        self.output_drop.drop_target_register(DND_FILES)

        self.output_drop.dnd_bind('<<Drop>>', self.drop_output_folder)



        # Sort Files button

        ttk.Button(

            self.main_frame,

            text="Sort Files",

            command=self.sort_files

        ).grid(row=10, column=0, columnspan=3, pady=10)



        # Result label

        self.result_label = tk.Label(

            self.main_frame,

            text="",

            font=("Arial", 9),

            bg="#f0f2f5",

            fg="#333333"

        )

        self.result_label.grid(row=11, column=0, columnspan=3, pady=2)



        # Credits frame

        self.credits_frame = tk.Frame(self.main_frame, bg="#ffffff", bd=1, relief="solid", padx=8, pady=5)

        self.credits_frame.grid(row=12, column=0, columnspan=3, pady=10, sticky="ew")



        tk.Label(

            self.credits_frame,

            text="Created by Bibhuti",

            font=("Arial", 9, "bold"),

            bg="#ffffff"

        ).pack(pady=1)



        fb_link = tk.Label(

            self.credits_frame,

            text="Facebook.com/bibhutithecoolboy",

            font=("Arial", 9),

            fg="#0066cc",

            bg="#ffffff",

            cursor="hand2"

        )

        fb_link.pack(pady=1)

        fb_link.bind("<Button-1>", lambda e: self.open_facebook())



        tk.Label(

            self.credits_frame,

            text="If you find this tool useful, please consider donating.",

            font=("Arial", 9),

            bg="#ffffff"

        ).pack(pady=1)



        tk.Label(

            self.credits_frame,

            text="If you need any custom tools, contact me via Facebook.",

            font=("Arial", 9),

            bg="#ffffff"

        ).pack(pady=1)



        # Donate Now button

        ttk.Button(

            self.credits_frame,

            text="Donate Now",

            command=self.show_qr_code

        ).pack(pady=5)



        # Configure grid weights

        self.main_frame.columnconfigure(0, weight=1)

        self.main_frame.columnconfigure(1, weight=1)

        self.main_frame.columnconfigure(2, weight=0)



        # Center the main window on the screen

        self.root.update_idletasks()

        width = self.root.winfo_width()

        height = self.root.winfo_height()

        screen_width = self.root.winfo_screenwidth()

        screen_height = self.root.winfo_screenheight()

        x = (screen_width // 2) - (width // 2)

        y = (screen_height // 2) - (height // 2)

        self.root.geometry(f"{width}x{height}+{x}+{y}")



    def select_excel_file(self):

        file_path = filedialog.askopenfilename(

            title="Select Excel/CSV File",

            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]

        )

        if file_path:

            self.entry_excel.delete(0, tk.END)

            self.entry_excel.insert(0, file_path)



    def select_image_folder(self):

        folder = filedialog.askdirectory(title="Select Image Folder")

        if folder:

            self.entry_image.delete(0, tk.END)

            self.entry_image.insert(0, folder)



    def select_output_folder(self):

        folder = filedialog.askdirectory(title="Select Output Folder")

        if folder:

            self.entry_output.delete(0, tk.END)

            self.entry_output.insert(0, folder)



    def sort_files(self):

        excel_file = self.entry_excel.get().strip()

        image_folder = self.entry_image.get().strip()

        output_folder = self.entry_output.get().strip()



        if not excel_file or not os.path.isfile(excel_file):

            messagebox.showerror("Error", "Please select a valid Excel/CSV file!")

            return



        if not image_folder or not os.path.isdir(image_folder):

            messagebox.showerror("Error", "Please select a valid Image folder!")

            return



        # If output folder is not provided, create a default one

        if not output_folder:

            output_folder = os.path.join(image_folder, "Sorted_Images")



        try:

            if not os.path.exists(output_folder):

                os.makedirs(output_folder)



            image_files_to_sort = []

            try:

                if excel_file.lower().endswith('.xlsx'):

                    df = pd.read_excel(excel_file, header=None, usecols=[0])

                    image_files_to_sort = [os.path.splitext(f)[0] for f in df[0].astype(str).tolist()]

                elif excel_file.lower().endswith('.csv'):

                    with open(excel_file, 'r', newline='') as f:

                        reader = csv.reader(f)

                        image_files_to_sort = [os.path.splitext(row[0])[0] for row in reader if row]

                else:

                    messagebox.showerror("Error", "Unsupported file type. Please select an .xlsx or .csv file.")

                    return

            except FileNotFoundError:

                messagebox.showerror("Error", f"The file was not found: {excel_file}")

                return

            except ModuleNotFoundError:

                 messagebox.showerror("Error", "The required library 'pandas' or 'openpyxl' is not installed. Please install them using: pip install pandas openpyxl")

                 return

            except Exception as e:

                messagebox.showerror("Error", f"Failed to read the spreadsheet file: {e}")

                return



            # Copy matching files

            copied = 0

            for file_in_folder in os.listdir(image_folder):

                base_name, _ = os.path.splitext(file_in_folder)

                if base_name in image_files_to_sort:

                    shutil.copy(os.path.join(image_folder, file_in_folder), 

                              os.path.join(output_folder, file_in_folder))

                    copied += 1



            self.result_label.config(text=f"Success! {copied} files copied successfully!", fg="#008800")

            messagebox.showinfo("Success", f"{copied} files copied successfully!")



        except Exception as e:

            self.result_label.config(text=f"Error: {str(e)}", fg="#cc0000")

            messagebox.showerror("Error", f"Failed to sort files: {str(e)}")

			

    def drop_excel_file(self, event):

        data = event.data.strip('{}')

        if os.path.isfile(data):

            self.entry_excel.delete(0, tk.END)

            self.entry_excel.insert(0, data)



    def drop_image_folder(self, event):

        data = event.data.strip('{}')

        if os.path.isdir(data):

            self.entry_image.delete(0, tk.END)

            self.entry_image.insert(0, data)



    def drop_output_folder(self, event):

        data = event.data.strip('{}')

        if os.path.isdir(data):

            self.entry_output.delete(0, tk.END)

            self.entry_output.insert(0, data)



    def open_facebook(self):

        webbrowser.open("https://facebook.com/bibhutithecoolboy")



    def show_qr_code(self):

        try:

            # Load QR code image, handling both script and EXE modes

            if hasattr(sys, '_MEIPASS'):

                qr_path = os.path.join(sys._MEIPASS, 'payment_qr_code.png')

            else:

                qr_path = os.path.join(os.path.dirname(__file__), 'payment_qr_code.png')



            if not os.path.exists(qr_path):

                raise FileNotFoundError("QR code image not found at: " + qr_path)



            qr_image = Image.open(qr_path)

            qr_photo = ImageTk.PhotoImage(qr_image)



            # Create a separate window for QR code

            qr_window = Toplevel(self.root)

            qr_window.title("Donate via QR Code")

            qr_window.resizable(False, False)

            qr_window.configure(bg="#f0f2f5")



            # QR code label

            qr_label = tk.Label(qr_window, image=qr_photo, bg="#f0f2f5")

            qr_label.image = qr_photo  # Keep a reference

            qr_label.pack(pady=10)



            # Close button

            ttk.Button(

                qr_window,

                text="Close",

                command=qr_window.destroy

            ).pack(pady=10)



            # Set window size based on original image size

            qr_window.update_idletasks()

            width = qr_image.width + 20

            height = qr_image.height + 80

            qr_window.geometry(f"{width}x{height}")



            # Center the QR window on the screen

            screen_width = qr_window.winfo_screenwidth()

            screen_height = qr_window.winfo_screenheight()

            x = (screen_width // 2) - (width // 2)

            y = (screen_height // 2) - (height // 2)

            qr_window.geometry(f"{width}x{height}+{x}+{y}")



        except Exception as e:

            messagebox.showerror("Error", f"Failed to load QR code image: {str(e)}\n\nPlease ensure 'payment_qr_code.png' is in the same folder as the script or EXE.")



if __name__ == "__main__":

    root = TkinterDnD.Tk()

    app = FileSorter(root)

    root.mainloop()

	

	
