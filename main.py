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
import pandas as pd
from pathlib import Path  # Import pathlib

def current_datetime_string():
    now = datetime.datetime.now()
    formatted_datetime = now.strftime("%Y-%m-%d-%H_%M_%S")
    return formatted_datetime

source_language = 'ch_sim'  # language code convention as defined in EasyOCR
source_folder = 'input'
source_folder = Path(source_folder)
current_time = current_datetime_string()
output_name_part = 'List of Fapiaos'
output_file = f'{output_name_part}}_{current_time}.xlsx'

regex_for_xiaopiao = r'小写[^0-9]*([0-9]+(?:\.[0.9]+)?)'  # a fairly straightforward way to extract SUM from fapiao; works good on ePDF and JPGs and poorly on paper scans
ocr_extensions_img = ['.jpg', '.jpeg', '.png']
pdf_extensions = ['.pdf']
all_extensions = ocr_extensions_img + pdf_extensions

def count_files_with_extensions(folder_path, extensions):
    if not os.path.exists(folder_path):
        return 0

    file_count = 0

    for extension in extensions:
        search_pattern = os.path.join(folder_path, f'*{extension}')
        matching_files = glob.glob(search_pattern)
        file_count += len(matching_files)
    return file_count

print(all_extensions)
print("Files in the folder:")
print(count_files_with_extensions(source_folder, all_extensions))

def extract_numbers_from_image(image_path, max_rotation_attempts=3):
    rotation_attempts = 0  # if SUM is not found, try to rotate the image (maybe the file is upside down)
    while rotation_attempts < max_rotation_attempts:
        print(f"Performing OCR on {image_path} (Attempt {rotation_attempts + 1})...")
        image = cv2.imread(image_path)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        reader = easyocr.Reader([source_language])
        results = reader.readtext(gray)
        #print(results)
        print(f"Searching for the SUM in {image_path}...")

        for result in results:
            text = result[1]  # Extract the text from the result
            match = re.search(regex_for_xiaopiao, text)
            if match:
                print(f"The sum is {match.group(1)} RMB")
                return match.group(1)  # Return the extracted numbers from the capturing group
        
        # Rotate the image clockwise by 90 degrees for the next attempt
        image = cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)
        cv2.imwrite(image_path, image)
        rotation_attempts += 1

def save_dict_to_excel(result_dict, output_file):
    data = {'Filename': list(result_dict.keys()), 'Values': list(result_dict.values())}
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)

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
            os.remove(jpg_path)
    elif image_path.suffix.lower() in (ocr_extensions_img):
        # Convert image_path to a string before passing to extract_numbers_from_image
        extracted_value = extract_numbers_from_image(str(image_path))
        if extracted_value is not None:
            try:
                result_dict[image_path.name] = float(extracted_value)
            except:
                result_dict[image_path.name] = extracted_value

print(result_dict)
save_dict_to_excel(result_dict, output_file)
