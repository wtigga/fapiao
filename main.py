import cv2
import easyocr
import re
import os
import fitz
from PIL import Image
import pandas as pd


source_language = 'ch_sim'
source_folder = 'input'  # Replace with your folder path
output_file = 'output.xlsx'

regex_for_xiaopiao = r'小写[^0-9]*([0-9]+(?:\.[0.9]+)?)'

def extract_numbers_from_image(image_path, max_rotation_attempts=3):
    rotation_attempts = 0
    while rotation_attempts < max_rotation_attempts:
        print(f"Performing OCR on {image_path} (Attempt {rotation_attempts + 1})...")
        image = cv2.imread(image_path)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        reader = easyocr.Reader([source_language])
        results = reader.readtext(gray)
        print(results)
        print(f"Searching for the SUM in {image_path}...")

        for result in results:
            text = result[1]  # Extract the text from the result
            match = re.search(r'小写[^0-9]*([0-9]+(?:\.[0-9]+)?)', text)
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

for filename in os.listdir(source_folder):
    file_path = os.path.join(source_folder, filename)

    if filename.lower().endswith('.pdf'):
        pdf_document = fitz.open(file_path)
        for page_number in range(pdf_document.page_count):
            page = pdf_document.load_page(page_number)
            
            # Set the DPI for rendering the page as an image (e.g., 300 DPI)
            dpi = 150
            image_list = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
            
            jpg_file = f"{filename}_page_{page_number + 1}.jpg"
            jpg_path = os.path.join(source_folder, jpg_file)
            
            img = Image.frombytes("RGB", [image_list.width, image_list.height], image_list.samples)
            img.save(jpg_path)
            
            extracted_value = extract_numbers_from_image(jpg_path)
            if extracted_value is not None:
                result_dict[jpg_file] = extracted_value
    elif filename.lower().endswith(('.jpg', '.png')):
        extracted_value = extract_numbers_from_image(file_path)
        if extracted_value is not None:
            result_dict[filename] = extracted_value

print(result_dict)
save_dict_to_excel(result_dict, output_file)
