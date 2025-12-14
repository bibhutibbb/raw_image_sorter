# File Sorter

A user-friendly application to efficiently sort all types of files from a source folder into a destination folder based on a list of filenames provided in an Excel or CSV file.

---

## Features

- **Flexible Input:** Select your list of target files using either an Excel (`.xlsx`) or a CSV (`.csv`) file.
- **Copy or Move:** Choose to either **Copy** the files to the destination or **Move** them permanently.
- **Sorts All Related Files:** If your spreadsheet lists `photo-001.jpg`, the app will find and process `photo-001.jpg`, `photo-001.xmp`, `photo-001.raw`, and any other file that shares the `photo-001` base name.
- **Smart Output Folder:** You can specify a custom output folder. If you don't, a `Sorted_Images` folder will be created inside the source image folder automatically.
- **Responsive UI:** The interface remains responsive during long operations thanks to background processing.
- **Progress Tracking:** A real-time progress bar and status text keep you informed of the sorting process.
- **Cancel Operation:** A "Cancel" button allows you to stop the file sorting process at any time.
- **Drag and Drop:** Easily drag and drop your spreadsheet, source folder, and output folder into the respective fields.

---

## How to Use

1.  **Run the application:**
    ```bash
    python raw_image_sorter_v3.py
    ```
2.  **Select Spreadsheet:** Click "Browse" or drag and drop your `.xlsx` or `.csv` file. The file should contain the base names of the files you want to sort in the first column (e.g., `photo-001`, `photo-002.jpg`, etc.).
3.  **Select Image Folder:** Click "Browse" or drag and drop the main source folder that contains all the files you want to sort through.
4.  **Select Output Folder (Optional):** Click "Browse" or drag and drop a folder where you want the sorted files to go. If left empty, a `Sorted_Images` folder will be created inside the Image Folder.
5.  **Choose Operation Mode:** Select either "Copy" or "Move" from the dropdown menu.
    - **Copy:** Makes a duplicate of the matching files in the output folder.
    - **Move:** Transfers the matching files from the source to the output folder.
6.  **Sort Files:** Click the "Sort Files" button to begin.
7.  **Monitor or Cancel:** Watch the progress bar. If you need to stop the process, click the "Cancel" button.

---

## Credits & Support

- **Created by Bibhuti**
- For custom tools or support, contact via Facebook: [facebook.com/bibhutithecoolboy](https://facebook.com/bibhutithecoolboy)

If you find this app helpful, please consider showing your support by donating.

<p align="center">
  <img src="payment_qr_code.png" alt="Donate QR Code" />
</p>