# pdf-table-to-excel-extraction
Given a single-page PDF, extract the table and output it into Excel format.

The top row of the Excel file will be the headers of the table. Each row in the table corresponds to one Excel row, and likewise for Excel columns.

* Pdf is first converted into an image using Pdf2image library.
* Image is then converted to grayscale.
* Median blur is done to remove noise.
* Used Tesseract to apply OCR and extract text from image and to localize each area of text in the image.
* Extracted the bounding box coordinates of the text region.
* Extracted PDF columns on the basis of bounding box 'top' coordinate and form column dataframe.
* Adjusted rows to its respective row number in column dataframe.
* Transposed the dataframe and set column names.
* Standardized date entries in YYYY/MM/DD format.
* Standardize numerical entries in ###.## format by removing commas.
* Write dataframe to excel file.
