from pytesseract import pytesseract, Output
from datetime import datetime
import pandas as pd
import pdf2image
import cv2
import os

current_path = os.path.abspath(__file__)
curr_dir = os.path.dirname(current_path)
parent_dir = os.path.dirname(curr_dir)
path = os.path.join(parent_dir, "data\\canopy_technical_test_input.")

pdf_file = f"{path}pdf"
img_file = f"{path}jpg"
txt_file = f"{path}txt"
excel_file = f"{path}xlsx"

# PDF to image, save image file
img = pdf2image.convert_from_path(pdf_file)
img[0].save(img_file, 'JPEG')

# load the example image and convert it to grayscale, Do median blur to remove noise & write the grayscale image to disk
image = cv2.imread(img_file)
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
gray = cv2.medianBlur(gray, 3)
cv2.imwrite(img_file, gray)

# load the image as a PIL/Pillow image, apply OCR, use Tesseract to localize each area of text in the input image
img = cv2.imread(img_file)
data = pytesseract.image_to_data(img, output_type=Output.DICT)

columns = ['text', 'conf', 'left', 'top', 'width', 'height']
word, conf, x, y, w, h, rows = [], [], [], [], [], [], []

# loop over each of the individual text localizations
for i in range(22, 242):
    # extract the OCR text itself along with the confidence of the text localization
    word = data["text"][i]
    conf = int(data["conf"][i])
    # extract the bounding box coordinates of the text region from the current result
    x = data["left"][i]
    y = data["top"][i]
    w = data["width"][i]
    h = data["height"][i]
    rows.append([word, conf, x, y, w, h])

# Dataframe with individual text localizations
boundaries = pd.DataFrame(rows, columns=columns)


# Get separate column based on bounding box 'left' coordinate
def get_col_startswith(tup):
    col = boundaries[boundaries.left.astype(str).str.startswith(tup)]
    return col


# Extract PDF columns on the basis of bounding box 'top' coordinate
# 'top' coordinates are decided from the output of boundaries
booking_date = get_col_startswith(('3', '4'))
txn_date = get_col_startswith('5')
booking_text = get_col_startswith(('6', '7', '8', '9', '10', '11', '12', '13'))
value_date = get_col_startswith(('14', '15', '16'))
debit = get_col_startswith('17')
credit = get_col_startswith(('18', '19', '20'))
balance = get_col_startswith(('21', '22'))
columns = [booking_date, txn_date, booking_text, value_date, debit, credit, balance]

# extract all row indexes by bounding box 'top' coordinate, then pair row boundary index sequentially
rows = boundaries.top.unique()
dict_row = {row: index for index, row in enumerate(rows)}

# Adjust rows to its respective row number in column dataframe
table_columns = []
for col in columns:
    # Initialize column empty list
    column = ['' for _ in range(len(dict_row))]
    # Loop through column elements
    for top, text in zip(col.top, col.text):
        column[dict_row[top]] = ' '.join((str(column[dict_row[top]]), text))
    table_columns.append(column[2:])

# Transpose dataframe
table_excel = pd.DataFrame(table_columns).T

# Setting dataframe column names
columns = ['Booking Date', 'Txn Date', 'Booking Text', 'Value Date', 'Debit', 'Credit', 'Balance']
table_excel.columns = columns


def format_date(dt):
    dt = datetime.strptime(dt.strip(), "%d.%m.%Y").strftime("%Y/%m/%d") if dt.strip() != '' else ''
    return dt


# standardize date entries in YYYY/MM/DD format
date_cols = [col for col in table_excel.columns if 'Date' in col]
table_excel[date_cols] = table_excel[date_cols].astype(str).applymap(lambda m: format_date(m))

# standardize numerical entries in ###.## format by removing commas
num_cols = [col for col in table_excel.columns if 'Date' not in col and 'Text' not in col]
table_excel[num_cols] = table_excel[num_cols].applymap(lambda ele: ele.replace(',', ''))

# Write to excel file
table_excel.to_excel(excel_file, index=False, index_label=False)
