# Created by Bibhuti (Facebook.com/bibhutithecoolboy) If you find it useful, please consider donating something.
# Contact me on Facebook to get the donation details.
# If you need any custom tools, feel free to contact me.

# File Sorter - Sorts files based on matching file names from excel file.
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
import threading

class FileSorter:
    __version__ = "3.0.0.0"
    def __init__(self, root):
        self.root = root
        self.root.title(f"File Sorter v{self.__version__}")
        self.root.configure(bg="#f0f2f5")
        self.root.resizable(True, True)
        self.cancel_event = threading.Event()
        try:
            # Determine base path for resources
            if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(base_path, "app_icon.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Error setting icon: {e}")
        self.style = ttk.Style()
        self.style.configure("TButton", font=("Arial", 10), padding=5)
        self.style.map("TButton", background=[("active", "#e0e0e0")])
        self.style.configure("TLabel", font=("Arial", 10), background="#f0f2f5")
        self.main_frame = tk.Frame(root, bg="#f0f2f5", padx=15, pady=15)
        self.main_frame.pack(fill="both", expand=True)
        tk.Label(self.main_frame, text="File Sorter", font=("Arial", 14, "bold"), bg="#f0f2f5", fg="#333333").grid(row=0, column=0, columnspan=3, pady=(0, 10))
        tk.Label(self.main_frame, text="Select Excel/CSV File or Drag and Drop:", font=("Arial", 9), bg="#f0f2f5").grid(row=1, column=0, sticky="w", pady=2)
        self.entry_excel = tk.Entry(self.main_frame, width=35, font=("Arial", 9))
        self.entry_excel.grid(row=2, column=0, columnspan=2, sticky="ew", padx=(0, 5))
        self.entry_excel.drop_target_register(DND_FILES)
        self.entry_excel.dnd_bind('<<Drop>>', self.drop_excel_file)
        ttk.Button(self.main_frame, text="Browse", command=self.select_excel_file).grid(row=2, column=2, sticky="w")
        self.excel_drop = tk.Label(self.main_frame, text="Drop Excel/CSV File Here", font=("Arial", 9, "italic"), bg="#e6f3ff", bd=2, relief="groove", height=1, width=35)
        self.excel_drop.grid(row=3, column=0, columnspan=2, pady=2, padx=(0, 5))
        self.excel_drop.drop_target_register(DND_FILES)
        self.excel_drop.dnd_bind('<<Drop>>', self.drop_excel_file)
        tk.Label(self.main_frame, text="Select Image Folder or Drag and Drop:", font=("Arial", 9), bg="#f0f2f5").grid(row=4, column=0, sticky="w", pady=2)
        self.entry_image = tk.Entry(self.main_frame, width=35, font=("Arial", 9))
        self.entry_image.grid(row=5, column=0, columnspan=2, sticky="ew", padx=(0, 5))
        self.entry_image.drop_target_register(DND_FILES)
        self.entry_image.dnd_bind('<<Drop>>', self.drop_image_folder)
        ttk.Button(self.main_frame, text="Browse", command=self.select_image_folder).grid(row=5, column=2, sticky="w")
        self.image_drop = tk.Label(self.main_frame, text="Drop Image Folder Here", font=("Arial", 9, "italic"), bg="#e6f3ff", bd=2, relief="groove", height=1, width=35)
        self.image_drop.grid(row=6, column=0, columnspan=2, pady=2, padx=(0, 5))
        self.image_drop.drop_target_register(DND_FILES)
        self.image_drop.dnd_bind('<<Drop>>', self.drop_image_folder)
        tk.Label(self.main_frame, text="Select Output Folder or Drag and Drop:", font=("Arial", 9), bg="#f0f2f5").grid(row=7, column=0, sticky="w", pady=2)
        self.entry_output = tk.Entry(self.main_frame, width=35, font=("Arial", 9))
        self.entry_output.grid(row=8, column=0, columnspan=2, sticky="ew", padx=(0, 5))
        self.entry_output.drop_target_register(DND_FILES)
        self.entry_output.dnd_bind('<<Drop>>', self.drop_output_folder)
        ttk.Button(self.main_frame, text="Browse", command=self.select_output_folder).grid(row=8, column=2, sticky="w")
        self.output_drop = tk.Label(self.main_frame, text="Drop Output Folder Here", font=("Arial", 9, "italic"), bg="#e6f3ff", bd=2, relief="groove", height=1, width=35)
        self.output_drop.grid(row=9, column=0, columnspan=2, pady=2, padx=(0, 5))
        self.output_drop.drop_target_register(DND_FILES)
        self.output_drop.dnd_bind('<<Drop>>', self.drop_output_folder)
        tk.Label(self.main_frame, text="Operation Mode:", font=("Arial", 9), bg="#f0f2f5").grid(row=10, column=0, sticky="w", pady=(5,0))
        self.operation_mode = ttk.Combobox(self.main_frame, values=["Copy", "Move"], state="readonly", font=("Arial", 9))
        self.operation_mode.set("Copy")
        self.operation_mode.grid(row=10, column=1, columnspan=2, sticky="ew", padx=(0, 5), pady=(5,0))
        self.sort_button = ttk.Button(self.main_frame, text="Sort Files", command=self.sort_files)
        self.sort_button.grid(row=11, column=0, columnspan=3, pady=10)
        self.progress_label = tk.Label(self.main_frame, text="", font=("Arial", 9), bg="#f0f2f5")
        self.progress_label.grid(row=12, column=0, columnspan=3, pady=2)
        self.progress_bar = ttk.Progressbar(self.main_frame, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.grid(row=13, column=0, columnspan=3, pady=5)
        self.cancel_button = ttk.Button(self.main_frame, text="Cancel", command=self.cancel_sorting)
        self.cancel_button.grid(row=14, column=0, columnspan=3, pady=5)
        self.cancel_button.grid_remove()
        self.result_label = tk.Label(self.main_frame, text="", font=("Arial", 9), bg="#f0f2f5", fg="#333333")
        self.result_label.grid(row=15, column=0, columnspan=3, pady=2)
        self.credits_frame = tk.Frame(self.main_frame, bg="#ffffff", bd=1, relief="solid", padx=8, pady=5)
        self.credits_frame.grid(row=16, column=0, columnspan=3, pady=10, sticky="ew")
        tk.Label(self.credits_frame, text="Created by Bibhuti", font=("Arial", 9, "bold"), bg="#ffffff").pack(pady=1)
        fb_link = tk.Label(self.credits_frame, text="Facebook.com/bibhutithecoolboy", font=("Arial", 9), fg="#0066cc", bg="#ffffff", cursor="hand2")
        fb_link.pack(pady=1)
        fb_link.bind("<Button-1>", lambda e: self.open_facebook())
        tk.Label(self.credits_frame, text="If you find this tool useful, please consider donating.", font=("Arial", 9), bg="#ffffff").pack(pady=1)
        tk.Label(self.credits_frame, text="If you need any custom tools, contact me via Facebook.", font=("Arial", 9), bg="#ffffff").pack(pady=1)
        ttk.Button(self.credits_frame, text="Donate Now", command=self.show_qr_code).pack(pady=5)
        self.main_frame.columnconfigure(0, weight=1)
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(title="Select Excel/CSV File", filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")])
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

    def cancel_sorting(self):
        self.cancel_event.set()

    def sort_files(self):
        excel_file = self.entry_excel.get().strip()
        image_folder = self.entry_image.get().strip()
        output_folder = self.entry_output.get().strip()
        mode = self.operation_mode.get()
        if not excel_file or not os.path.isfile(excel_file):
            messagebox.showerror("Error", "Please select a valid Excel/CSV file!")
            return
        if not image_folder or not os.path.isdir(image_folder):
            messagebox.showerror("Error", "Please select a valid Image folder!")
            return
        if not output_folder:
            output_folder = os.path.join(image_folder, "Sorted_Images")
        self.sort_button.config(state="disabled")
        self.cancel_button.grid()
        self.progress_bar["value"] = 0
        self.progress_label['text'] = ""
        self.result_label['text'] = ""
        self.cancel_event.clear()
        sorter_thread = threading.Thread(
            target=self._execute_sorting,
            args=(excel_file, image_folder, output_folder, mode)
        )
        sorter_thread.start()

    def _execute_sorting(self, excel_file, image_folder, output_folder, mode):
        try:
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            image_files_to_sort = []
            try:
                if excel_file.lower().endswith('.xlsx'):
                    df = pd.read_excel(excel_file, header=None, usecols=[0])
                    image_files_to_sort = [os.path.splitext(f)[0] for f in df[0].astype(str).tolist()]
                elif excel_file.lower().endswith('.csv'):
                    with open(excel_file, 'r', newline='', encoding='utf-8') as f:
                        reader = csv.reader(f)
                        image_files_to_sort = [os.path.splitext(row[0])[0] for row in reader if row]
            except Exception as e:
                self.root.after_idle(messagebox.showerror, "Error", f"Failed to read the spreadsheet file: {e}")
                self.root.after_idle(self._sorting_complete, "Error", 0, mode.lower())
                return
            image_files_set = set(image_files_to_sort)
            all_files_in_folder = os.listdir(image_folder)
            files_to_process = [f for f in all_files_in_folder if os.path.splitext(f)[0] in image_files_set]
            total_files = len(files_to_process)
            self.root.after_idle(self._update_progress, 0, total_files, mode.lower())
            processed_count = 0
            for filename in files_to_process:
                if self.cancel_event.is_set():
                    break
                source_path = os.path.join(image_folder, filename)
                dest_path = os.path.join(output_folder, filename)
                try:
                    if mode == "Copy":
                        shutil.copy(source_path, dest_path)
                    elif mode == "Move":
                        shutil.move(source_path, dest_path)
                    processed_count += 1
                except Exception as e:
                    print(f"Could not process file {filename}: {e}")
                self.root.after_idle(self._update_progress, processed_count, total_files, mode.lower())
            status = "Cancelled" if self.cancel_event.is_set() else "Success"
            self.root.after_idle(self._sorting_complete, status, processed_count, mode.lower())
        except Exception as e:
            self.root.after_idle(self._sorting_complete, f"Error: {e}", 0, mode.lower())

    def _update_progress(self, processed_count, total_count, mode_verb):
        action_text = "moved" if mode_verb == "move" else "copied"
        if total_count > 0:
            percentage = (processed_count / total_count) * 100
            self.progress_bar["value"] = percentage
            self.progress_label['text'] = f"{action_text.capitalize()} {processed_count} / {total_count} files"
        else:
            self.progress_bar["value"] = 0
            self.progress_label['text'] = "No matching files to process."

    def _sorting_complete(self, status, processed_count, mode_verb):
        action_text = "moved" if mode_verb == "move" else "copied"
        self.sort_button.config(state="normal")
        self.cancel_button.grid_remove()
        self.progress_bar["value"] = 0
        if "Error" in status:
            self.result_label.config(text=status, fg="#cc0000")
            messagebox.showerror("Error", f"Failed to sort files: {status}")
        elif status == "Cancelled":
            self.result_label.config(text=f"Operation cancelled after {action_text} {processed_count} files.", fg="#ff8c00")
        else:
            self.result_label.config(text=f"Success! {processed_count} files {action_text} successfully!", fg="#008800")
            messagebox.showinfo("Success", f"{processed_count} files {action_text} successfully!")
        self.progress_label['text'] = ""

    def open_facebook(self):
        webbrowser.open("https://facebook.com/bibhutithecoolboy")

    def show_qr_code(self):
        try:
            if hasattr(sys, '_MEIPASS'):
                qr_path = os.path.join(sys._MEIPASS, 'payment_qr_code.png')
            else:
                qr_path = os.path.join(os.path.dirname(__file__), 'payment_qr_code.png')
            if not os.path.exists(qr_path):
                raise FileNotFoundError("QR code image not found")
            qr_image = Image.open(qr_path)
            qr_photo = ImageTk.PhotoImage(qr_image)
            qr_window = Toplevel(self.root)
            qr_window.title("Donate via QR Code")
            qr_window.resizable(False, False)
            qr_window.configure(bg="#f0f2f5")
            qr_label = tk.Label(qr_window, image=qr_photo, bg="#f0f2f5")
            qr_label.image = qr_photo
            qr_label.pack(pady=10)
            ttk.Button(qr_window, text="Close", command=qr_window.destroy).pack(pady=10)
            qr_window.update_idletasks()
            width = qr_image.width + 20
            height = qr_image.height + 80
            x = (qr_window.winfo_screenwidth() // 2) - (width // 2)
            y = (qr_window.winfo_screenheight() // 2) - (height // 2)
            qr_window.geometry(f"{width}x{height}+{x}+{y}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load QR code image: {e}")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = FileSorter(root)
    root.mainloop()
