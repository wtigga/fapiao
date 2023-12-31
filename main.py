import cv2
import easyocr
import re
import datetime
import fitz
from PIL import Image
import os
import PIL
import uuid
import xlsxwriter
import time
import tkinter as tk
import threading
from tkinter import filedialog, messagebox, ttk
import webbrowser
import shutil
import subprocess

PIL.Image.ANTIALIAS = (
    PIL.Image.LANCZOS
)


#  this is a workaround for outdated EasyOCR library
# https://github.com/JaidedAI/EasyOCR/issues/1077

def current_datetime_string():
    # for a timestamp
    now = datetime.datetime.now()
    formatted_datetime = now.strftime("%Y-%m-%d-%H_%M_%S")
    return formatted_datetime


script_name = "Fapiao OCR"
script_version = "0.3"
script_date = "2023-10-31"  # last update
script_title = f"{script_name}, ver.{script_version}"
source_language = (
    "ch_sim"  # language code for Chinese Simplified convention as defined in EasyOCR
)
source_folder = ""
current_time = current_datetime_string()
output_name_part = "List of Fapiaos"
output_file = f"{output_name_part}_{current_time}.xlsx"
regex_text = "小写"
value_regex = r"^\d+(\.\d+)?$"  # Regular expression to match floats or integers
regex_for_xiaopiao = r".*小.*写.*?(\d+(?:[.,]\d+)?)"  # a fairly straightforward way to extract SUM from fapiao; works good on ePDF and JPGs and poorly on paper scans
ocr_extensions_img = [".jpg", ".jpeg", ".png"]
pdf_extensions = [".pdf"]
all_extensions = ocr_extensions_img + pdf_extensions

progress_bar_total = 100
progress_bar_current = 0

sum_not_found_files = []


def get_files_in_folder_with_extensions(folder_path, allowed_extensions):
    # Initialize an empty list to store the matching file names
    matching_files = []

    # Check if the folder exists
    if os.path.exists(folder_path):
        # List all files and directories in the folder
        for filename in os.listdir(folder_path):
            # Check if the item is a file and has the allowed extensions
            if os.path.isfile(
                    os.path.join(folder_path, filename)
            ) and filename.lower().endswith(tuple(allowed_extensions)):
                matching_files.append(filename)

    return matching_files


def find_closest_value_on_same_y(results, target_text, value_regex):
    # For cases when the value is placed far away from the 小写 but on the same axis
    # Finds the closest value
    target_coords = None
    matching_values = []

    # Find the target text's coordinates
    for coords, text, _ in results:
        if re.search(target_text, text):
            target_coords = coords
            break

    # If the target text was found
    if target_coords:
        target_x1, target_x2, target_y1, target_y2 = (
            min(target_coords, key=lambda x: x[0])[0],
            max(target_coords, key=lambda x: x[0])[0],
            min(target_coords, key=lambda x: x[1])[1],
            max(target_coords, key=lambda x: x[1])[1],
        )

        # Find matching values on the same Y-axis
        for coords, text, _ in results:
            if re.match(value_regex, text):
                x1, x2, y1, y2 = (
                    min(coords, key=lambda x: x[0])[0],
                    max(coords, key=lambda x: x[0])[0],
                    min(coords, key=lambda x: x[1])[1],
                    max(coords, key=lambda x: x[1])[1],
                )
                # Check if the Y-axis position is within a margin of 3-5 pixels
                if (
                        target_y1 - 5 <= y1 <= target_y2 + 5
                        and target_y1 - 5 <= y2 <= target_y2 + 5
                ):
                    matching_values.append((coords, text, x1))

        # If there are matching values, return the one closest by X-axis
        if matching_values:
            matching_values.sort(key=lambda x: abs(x[2] - target_x1))
            return matching_values[0][1]

    return "0"


def extract_numbers_from_image(image_path):
    file_name_for_report = os.path.basename(image_path)
    max_rotation_attempts = 3
    global sum_not_found_files
    # Create a 'temp' subfolder if it doesn't exist
    current_directory = os.getcwd()
    temp_folder_short_name = "temp"
    temp_folder = os.path.join(current_directory, temp_folder_short_name)
    try:
        os.makedirs(temp_folder)
    except:
        print("Folder already exist")

    if not os.path.exists(image_path):
        print(f"File not found: {image_path}")
        return None

    # Check if the filename contains non-ASCII characters
    if not all(ord(char) < 128 for char in image_path):
        # Generate a unique name for the temporary copy
        temp_filename = f"{uuid.uuid4()}.png"
        temp_path = os.path.join(temp_folder, temp_filename)

        # Create a temporary copy of the image with a unique name
        shutil.copy(image_path, temp_path)

        # Set image_path to the temporary copy for OCR
        image_path = temp_path

    rotation_attempts = 0
    # rotate a few times if no sum is found on the first try - what if the image is upside down?
    while rotation_attempts < max_rotation_attempts:
        print(f"Performing OCR on {image_path} (Attempt {rotation_attempts + 1})...")

        image = cv2.imread(image_path)

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
        # print(results)

        print(f"Searching for the SUM in {str(os.path.basename(image_path))}...")

        for result in results:
            text = result[1]  # Extract the text from the result
            match = re.search(regex_for_xiaopiao, text)
            if match:
                print(f"The sum is {match.group(1)} RMB")
                return match.group(
                    1
                )  # Return the extracted numbers from the capturing group

        value = find_closest_value_on_same_y(results, regex_text, value_regex)
        if float(value) > 0:
            return value

        # Rotate the image clockwise by 90 degrees for the next attempt
        image = cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)
        cv2.imwrite(str(image_path), image)
        rotation_attempts += 1

    # If all attempts fail, return None
    sum_not_found_files.append(file_name_for_report)
    print(f"No sum found in {file_name_for_report}")
    return None


def save_to_xlsx(data, filename):
    if not data:
        print("No sums in fapiaos found, nothing to save")
    else:
        # Create a new XLSX workbook and add a worksheet
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        wrap_format = workbook.add_format({"text_wrap": True})

        # Define the header format with bold text
        header_format = workbook.add_format({"bold": True})

        # Set column widths and add the headers
        worksheet.set_column("A:A", 70)  # Adjust column width as needed
        worksheet.set_column("B:B", 10)  # Adjust column width as needed
        worksheet.write("A1", "Filename", header_format)
        worksheet.write("B1", "Sum", header_format)
        worksheet.freeze_panes(
            1, 0
        )  # 1 is the first row (zero-based), 0 is the first column (zero-based)

        # Define a format for the currency symbol
        currency_format = workbook.add_format({"num_format": "¥#,##0.00"})
        # currency_format = workbook.add_format({'num_format': '¥#,##0'})

        # Start writing data from row 2
        row = 0  # Start writing from the first row (0-based index)

        for filename, value in data.items():
            row += 1
            worksheet.write(row, 0, filename, wrap_format)
            worksheet.write(row, 1, value, currency_format)

        # Create a format for text wrapping
        wrap_format = workbook.add_format({"text_wrap": True})

        # Apply automatic line breaking to all columns
        worksheet.set_row(
            0, None, wrap_format
        )  # Set row 0 (header row) to use text wrapping

        # Close the workbook
        workbook.close()


def sum_dict_values(input_dict):
    total = 0
    for value in input_dict.values():
        if isinstance(value, (int, float)):
            total += value
    return total


def fapiao_ocr():
    global progress_bar_current
    result_dict = {}

    for file_name in os.listdir(source_folder):
        file_path = os.path.join(source_folder, file_name)
        file_extension = os.path.splitext(file_name)[1].lower()
        if file_name.lower().endswith(".pdf"):
            pdf_path = os.path.join(source_folder, file_name)
            pdf_document = fitz.open(pdf_path)
            for page_number in range(pdf_document.page_count):
                page = pdf_document.load_page(page_number)

                dpi = 150  # too large files takes longer to scan, 150 is enough
                image_list = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72))

                jpg_file = f"{file_name}_page_{page_number + 1}.jpg"
                jpg_path = os.path.join(source_folder, jpg_file)
                img = Image.frombytes(
                    "RGB", [image_list.width, image_list.height], image_list.samples
                )
                img.save(jpg_path)

                # Convert jpg_path to a string before passing to cv2.imread
                extracted_value = extract_numbers_from_image(jpg_path)
                progress_bar_current += 1
                update_progress_bar()
                if extracted_value is not None:
                    try:
                        result_dict[os.path.basename(file_name)] = float(
                            extracted_value
                        )
                    except:
                        result_dict[os.path.basename(file_name)] = extracted_value
                try:
                    os.remove(jpg_path)
                except:
                    print("Temp file wasn't removed.")
        elif file_extension in ocr_extensions_img:
            # Convert image_path to a string before passing to extract_numbers_from_image
            extracted_value = extract_numbers_from_image(file_path)
            progress_bar_current += 1
            update_progress_bar()
            if extracted_value is not None:
                try:
                    result_dict[file_name] = float(extracted_value)
                except:
                    result_dict[file_name] = extracted_value
    return result_dict


# GUI #


def disable_all_buttons():
    run_button.config(state=tk.DISABLED)
    browse_button.config(state=tk.DISABLED)


def enable_all_buttons():
    run_button.config(state=tk.NORMAL)
    browse_button.config(state=tk.NORMAL)


# Function to be executed when the "RUN" button is clicked
def run_script():
    global sum_not_found_files
    # Record the start time
    disable_all_buttons()

    def main_logic():
        global sum_not_found_files
        print("Running OCR on files...")
        try:
            start_time = time.time()
            result = fapiao_ocr()
            if not result:
                messagebox.showerror(
                    "Nothing found", "No sums found in source files, try another folder"
                )
                enable_all_buttons()
            else:
                print("Saving to XLSX...")
                save_to_xlsx(result, output_file)
                # You can place your script code here
                print("Done")
                enable_all_buttons()

                end_time = time.time()
                total_time = end_time - start_time
                print(f"Total run time: {total_time:.2f} seconds")
                fapiao_sum = round(sum_dict_values(result), 2)
                messagebox.showinfo(
                    "Complete",
                    f"Report has been saved, files processed: {number_of_files}."
                    f"\n\nTotal sum: {fapiao_sum} RMB."
                    f"\n\nSee the report in {output_file} in the same folder as the script."
                    f"\n\nIt took just {int(total_time)} seconds!"
                )
                try:
                    subprocess.Popen([output_file], shell=True)
                except Exception as e:
                    print(f"An error occurred: {e}")
                if not sum_not_found_files:
                    pass
                else:
                    messagebox.showinfo(
                        f"Files without sums: {len(sum_not_found_files)}",
                        f"\nNo SUM found in {len(sum_not_found_files)} file(s):\n"
                        f"\n{file_list_to_string(sum_not_found_files)}"
                    )
                sum_not_found_files = []  # reset the list of files where nothing was found
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


# Create the main window
root = tk.Tk()
root.geometry("300x150")
root.title(script_title)

# Create the "RUN" button and associate it with the run_script function
run_button = tk.Button(root, text="RUN", command=run_script)
run_button.grid(row=0, column=1, ipadx=10, ipady=10, padx=10, pady=10)

number_of_files = 0


def file_list_to_string(incoming_list):
    incoming_list = "\n".join(incoming_list)
    return incoming_list


def browse_folder():
    global progress_bar_total
    # browse_button
    disable_all_buttons()

    def main_logic():
        global number_of_files
        global source_folder
        global progress_bar_total
        source_folder = filedialog.askdirectory()
        list_of_files = get_files_in_folder_with_extensions(
            source_folder, all_extensions
        )
        progress_bar_total = len(list_of_files)
        set_total_length(progress_bar_total)
        nice_list_of_files = file_list_to_string(list_of_files)
        source_folder_var.set(source_folder)
        number_of_files = len(list_of_files)
        if not list_of_files:
            messagebox.showerror(
                f"No files found",
                f"The selected folder doesn't contain any PDF, JPG, JPEG, or PNG files to scan. Please select another folder.",
            )
        elif number_of_files > 100:
            messagebox.showinfo(
                f"Too many files: {number_of_files}",
                f"There's a lot of images in this folder, scanning will take a while. The list of files:"
                f"\n\n{nice_list_of_files}."
                f"\n\nThis will take about {number_of_files * 3} seconds to process."
                f"\n\nIf it's OK, press 'RUN' button.",
            )
        else:
            messagebox.showinfo(
                f"Files found: {number_of_files}",
                f"Following files will be scanned for Fapiao information:\n\n{nice_list_of_files}"
                f"\n\nThis will take about {number_of_files * 3} seconds to process."
                f"\n\nIf it's OK, press 'RUN' button.",
            )
        enable_all_buttons()

    main_thread = threading.Thread(target=main_logic)
    main_thread.start()


# Create browse button for file folder with fapiaos
source_folder_var = tk.StringVar()
browse_button = ttk.Button(
    root, text="Browse folder with Fapiaos", command=browse_folder
)
browse_button.grid(row=0, column=0, ipadx=10, ipady=10, padx=10, pady=10, sticky="w")


# Text in the bottom
def open_url(url):
    webbrowser.open(url)


about_label = tk.Label(
    root,
    text=f"github.com/wtigga/fapiao\n"
         f"{script_date}",
    fg="blue",
    cursor="hand2",
    justify="left",
)
about_label.bind(
    "<Button-1>", lambda event: open_url("https://github.com/wtigga/fapiao")
)
about_label.grid(row=31, column=0, sticky="w", padx=10, pady=0)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
progress_bar.grid(row=32, column=0, columnspan=2, padx=10, pady=10)


def set_total_length(new_total_length):
    global progress_bar_total
    progress_bar_total = new_total_length
    progress_bar['maximum'] = new_total_length


def update_progress_bar():
    # Update the progress bar's value based on progress_bar_current
    progress_bar['value'] = progress_bar_current
    root.update_idletasks()  # Update the GUI to reflect the changes


# Start the GUI main loop


root.mainloop()
