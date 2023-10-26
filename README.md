# Fapiao OCR - Invoice Information Extraction

![image](https://github.com/wtigga/fapiao/assets/7037184/460b6100-2dab-4322-a768-be70644497d6)

## 1. What It Does

The "Fapiao OCR" script is designed to automate the extraction of invoice (发票 fāpiào) information from image files and PDFs. It utilizes Optical Character Recognition (OCR) techniques to recognize and extract numerical values from the specific place in the documents. 

The main functionalities of the script include:
* Offline work, no internet connection required, data stays on your PC
* Scanning a specified folder for invoice files (in PDF, JPG, JPEG, or PNG formats).
* Extracting numeric values from these files using OCR.
* Saving the extracted data to an Excel (XLSX) file for further analysis.

Key Features:
* GUI, user friendly interface, no need to input commands in a console
* Multiple file formats are supported, making it flexible for different data sources.
* The script is equipped to handle rotation attempts for better OCR results.
* It provides a graphical user interface (GUI) for easy interaction with the user.

## 2. Limitations
While the "Fapiao OCR" script is a helpful tool for automating invoice data extraction, it has certain limitations:

### OCR Accuracy
OCR accuracy depends on the quality and clarity of the source documents. Handwritten, heavily stylized, or low-resolution text may result in inaccurate extractions.

### Language Support
The script is designed for invoices in Chinese Simplified (ch_sim) issued in Mainland China. It will not work with other documents.

### Handling Non-ASCII Characters
The script attempts to handle filenames with non-ASCII characters, but there may still be limitations in certain cases. In other words, files with non-English characters in name or path might cause problems.

### PDF Extraction
For PDF files, the script extracts images from each page, which can be time-consuming and may not work well with extremely large PDFs.

### File Types
The script focuses on PDFs, JPGs, JPEGs, and PNGs. It does not support other invoice file formats like Word documents or Excel spreadsheets.

### Large distributive
It uses a lot of 3rd party libraries, such as PyTorch, thus the size of the distributive is quite big large.

# 3. Libraries Used

The "Fapiao OCR" script relies on the following libraries and modules:

* Python: The script is written in Python, a versatile and widely-used programming language.

* OpenCV (cv2): OpenCV is used for image processing and manipulation.

* PyMuPDF (fitz): PyMuPDF is used for extracting images from PDF files.

* EasyOCR: This library is used for Optical Character Recognition.

* xlsxwriter: xlsxwriter is used for creating XLSX files to save the extracted data.

* Tkinter: Tkinter is used to create the graphical user interface (GUI) for user interaction.

* threading: Threading is used for running the main logic in a separate thread to avoid GUI freezing.

* webbrowser: The webbrowser module is used to open URL "About" URL (leading to this Git repository)

# 4. How to use

1. Run the script
2. Click 'Browse' to select the folder with Fapiaos (in PDF, JPG, JPEG, or PNG)
3. Click 'RUN'
4. The report will be stored in the same folder where you run the script from.

Sample report:

![image](https://github.com/wtigga/fapiao/assets/7037184/a904e687-586c-4f89-a0d1-b25a6baf26fe)
