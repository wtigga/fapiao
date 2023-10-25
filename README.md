This is a simple script to simplify the mundane process of getting all the sums from a multiple PDF or JPG files with 'Fapiao' - receipts issued in Mainland China.
Submitting 'fapiaos' is a choir that many foreigners who live in China had to do on a monthly basis. This program will extract all the sums from all the fapiaos in a folder and put them in a convenient Excel file.

It is written in Python and uses OCR techniques.
Only compatible with Chinese fapiaos - in PDF, JPG, PNG. It was designed to work with Electronic fapiaos, but it might work with scanned copies or photographs.

In order to get EasyOCR work in 2023, I had to modify it's easyocr\utils.py, lines 574 and 576, replacing ANTIALIAS with cv2.INTER_LANCZOS4:

The full updated function look like this now:

def compute_ratio_and_resize(img,width,height,model_height):
    '''
    Calculate ratio and resize correctly for both horizontal text
    and vertical case
    '''
    ratio = width/height
    if ratio<1.0:
        ratio = calculate_ratio(width,height)
        img = cv2.resize(img,(model_height,int(model_height*ratio)), interpolation=cv2.INTER_LANCZOS4)
    else:
        img = cv2.resize(img,(int(model_height*ratio),model_height),interpolation=cv2.INTER_LANCZOS4)
    return img,ratio


    This is due to EasyOCR outdated code which doesn't work with modern PIL.