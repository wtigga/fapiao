import cv2
import easyocr
import re
import os
import datetime
import fitz
from PIL import Image
import os
import PIL
import glob
PIL.Image.ANTIALIAS = PIL.Image.LANCZOS   #  this is a workaround for outdated EasyOCR library
# https://github.com/JaidedAI/EasyOCR/issues/1077
from pathlib import Path  # Import pathlib
from openpyxl import Workbook
import uuid
import xlsxwriter
import time
import tkinter as tk
import threading
from tkinter import filedialog, messagebox, ttk
import webbrowser




def current_datetime_string():
    # for a timestamp
    now = datetime.datetime.now()
    formatted_datetime = now.strftime("%Y-%m-%d-%H_%M_%S")
    return formatted_datetime

script_name = 'Fapiao OCR'
script_version = '0.1'
script_title = f"{script_name}, ver.{script_version}"
source_language = 'ch_sim'  # language code convention as defined in EasyOCR
source_folder = 'input'
source_folder = Path(source_folder)
current_time = current_datetime_string()
output_name_part = 'List of Fapiaos'
output_file = f'{output_name_part}_{current_time}.xlsx'
regex_for_xiaopiao = r'.*小.*写.*?(\d+(?:[.,]\d+)?)' # a fairly straightforward way to extract SUM from fapiao; works good on ePDF and JPGs and poorly on paper scans
ocr_extensions_img = ['.jpg', '.jpeg', '.png']
pdf_extensions = ['.pdf']
all_extensions = ocr_extensions_img + pdf_extensions

def get_files_in_folder_with_extensions(folder_path, allowed_extensions):
    # Initialize an empty list to store the matching file names
    matching_files = []

    # Check if the folder exists
    if os.path.exists(folder_path):
        # List all files and directories in the folder
        for filename in os.listdir(folder_path):
            # Check if the item is a file and has the allowed extensions
            if os.path.isfile(os.path.join(folder_path, filename)) and filename.lower().endswith(tuple(allowed_extensions)):
                matching_files.append(filename)

    return matching_files

def extract_numbers_from_image(image_path, max_rotation_attempts=3):
    # extracting
    image_path = Path(image_path)
    
    # Create a 'temp' subfolder if it doesn't exist
    temp_folder = Path("temp")
    temp_folder.mkdir(parents=True, exist_ok=True)
    
    if not image_path.is_file():
        print(f"File not found: {str(image_path)}")
        return None

    # Check if the filename contains non-ASCII characters
    if not all(ord(char) < 128 for char in image_path.name):
        # Generate a unique name for the temporary copy
        temp_filename = f"{uuid.uuid4()}.png"
        temp_path = temp_folder / temp_filename
        
        # Create a temporary copy of the image with a unique name
        image_path.rename(temp_path)
        
        # Set image_path to the temporary copy for OCR
        image_path = temp_path

    rotation_attempts = 0
    # rotate a few times if no sum is found on the first try - what if the image is upside down?
    while rotation_attempts < max_rotation_attempts:
        print(f"Performing OCR on {str(image_path)} (Attempt {rotation_attempts + 1})...")


        image = cv2.imread(str(image_path))

        # Get the dimensions of the image
        image_height, image_width, _ = image.shape
        # Calculate the coordinates for the bottom right quarter
        x_mid = image_width // 2
        y_mid = image_height // 2
        # Crop the bottom right quarter
        bottom_right_quarter = image[y_mid:, x_mid:, :]
        # Convert the cropped region to grayscale
        gray = cv2.cvtColor(bottom_right_quarter, cv2.COLOR_BGR2GRAY)

        reader = easyocr.Reader([source_language])
        results = reader.readtext(gray)


        print(f"Searching for the SUM in {str(image_path)}...")

        for result in results:
            text = result[1]  # Extract the text from the result
            match = re.search(regex_for_xiaopiao, text)
            if match:
                print(f"The sum is {match.group(1)} RMB")
                return match.group(1)  # Return the extracted numbers from the capturing group
        
        # Rotate the image clockwise by 90 degrees for the next attempt
        image = cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)
        cv2.imwrite(str(image_path), image)
        rotation_attempts += 1
    
    # If all attempts fail, return None
    return None


def save_to_xlsx(data, filename):
    if not data:
        print("No sums in fapiaos found, nothing to save")
    else:
        # Create a new XLSX workbook and add a worksheet
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        wrap_format = workbook.add_format({'text_wrap': True})

        # Define the header format with bold text
        header_format = workbook.add_format({'bold': True})

        # Set column widths and add the headers
        worksheet.set_column('A:A', 70)  # Adjust column width as needed
        worksheet.set_column('B:B', 10)  # Adjust column width as needed
        worksheet.write('A1', 'Filename', header_format)
        worksheet.write('B1', 'Sum', header_format)
        worksheet.freeze_panes(1, 0)  # 1 is the first row (zero-based), 0 is the first column (zero-based)

        # Define a format for the currency symbol
        currency_format = workbook.add_format({'num_format': '¥#,##0'})

        # Start writing data from row 2
        row = 0  # Start writing from the first row (0-based index)

        for filename, value in data.items():
            row += 1
            worksheet.write(row, 0, filename, wrap_format)
            worksheet.write(row, 1, value, currency_format)

        # Create a format for text wrapping
        wrap_format = workbook.add_format({'text_wrap': True})

        # Apply automatic line breaking to all columns
        worksheet.set_row(0, None, wrap_format)  # Set row 0 (header row) to use text wrapping

        # Close the workbook
        workbook.close()

def fapiao_ocr():

    result_dict = {}

    for image_path in source_folder.glob('*.*'):
        if image_path.suffix.lower() == '.pdf':
            pdf_document = fitz.open(image_path)
            for page_number in range(pdf_document.page_count):
                page = pdf_document.load_page(page_number)
                
                dpi = 150  # too large files takes longer to scan, 150 is enough
                image_list = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
                
                jpg_file = f"{image_path.stem}_page_{page_number + 1}.jpg"
                jpg_path = source_folder / jpg_file
                img = Image.frombytes("RGB", [image_list.width, image_list.height], image_list.samples)
                img.save(jpg_path)
                
                # Convert jpg_path to a string before passing to cv2.imread
                extracted_value = extract_numbers_from_image(str(jpg_path))
                if extracted_value is not None:
                    try:
                        result_dict[image_path.name] = float(extracted_value)
                    except:
                        result_dict[image_path.name] = extracted_value
                try:
                    os.remove(jpg_path)
                except:
                    print("Temp file wasn't removed.")
        elif image_path.suffix.lower() in (ocr_extensions_img):
            # Convert image_path to a string before passing to extract_numbers_from_image
            extracted_value = extract_numbers_from_image(str(image_path))
            if extracted_value is not None:
                try:
                    result_dict[image_path.name] = float(extracted_value)
                except:
                    result_dict[image_path.name] = extracted_value
    return result_dict




# Print the total run time
#



# GUI #

def disable_all_buttons():
    run_button.config(state=tk.DISABLED)
    browse_button.config(state=tk.DISABLED)


def enable_all_buttons():
    run_button.config(state=tk.NORMAL)
    browse_button.config(state=tk.NORMAL)

# Function to be executed when the "RUN" button is clicked
def run_script():
    # Record the start time
    start_time = time.time()
    disable_all_buttons()
    
    def main_logic():
        print("Running OCR on files...")
        try:
            result = fapiao_ocr()
            if not result:
                messagebox.showinfo("Nothing found", f"No sums found in source files, try another folder")
                enable_all_buttons()
            else:
                print("Saving to XLSX...")
                save_to_xlsx(result, output_file)
                # You can place your script code here
                print("Done")
                enable_all_buttons()
                messagebox.showinfo("Complete", f"Report has been saved, files processed: {number_of_files}, report is in the file {output_file} next to this script.")
        except Exception as exp:
            # Show popup window with error message
            messagebox.showerror("Error", str(exp))
            enable_all_buttons()
    try:
        main_thread = threading.Thread(target=main_logic)
        main_thread.start()
    except Exception as exp:
        # Show popup window with error message
        messagebox.showerror("Error", str(exp))
    # Record the end time
    end_time = time.time()
    # Calculate the total run time
    total_time = end_time - start_time
    print(f"Total run time: {total_time:.2f} seconds") 

   

# Create the main window
root = tk.Tk()
root.geometry("280x100")
root.title(script_title)

# Create the "RUN" button and associate it with the run_script function
run_button = tk.Button(root, text="RUN", command=run_script)
run_button.grid(row=0, column=1, ipadx=10, ipady=10, padx=10, pady=10)

number_of_files = 0

def browse_folder():
    # browse_button
    disable_all_buttons()

    def main_logic():
        global number_of_files
        global source_folder
        source_folder = filedialog.askdirectory()
        list_of_files = get_files_in_folder_with_extensions(source_folder, all_extensions)
        print(list_of_files)
        source_folder = Path(source_folder)
        source_folder_var.set(source_folder)
        number_of_files = len(list_of_files)
        enable_all_buttons()

    main_thread = threading.Thread(target=main_logic)
    main_thread.start()

# Create browse button for file folder with fapiaos
source_folder_var = tk.StringVar()
browse_button = ttk.Button(root, text="Browse folder with Fapiaos", command=browse_folder)
browse_button.grid(row=0, column=0, ipadx=10, ipady=10, padx=10, pady=10, sticky='w')

# Text in the bottom
def open_url(url):
    webbrowser.open(url)

about_label = tk.Label(root, text="github.com/wtigga/fapiao\nVladimir Zhdanov", fg="blue", cursor="hand2",
                       justify="left")
about_label.bind("<Button-1>",
                 lambda event: open_url("https://github.com/wtigga/fapiao"))
about_label.grid(row=31, column=0, sticky='w', padx=10, pady=0)

# Start the GUI main loop
root.mainloop()