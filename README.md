# dynamic-pdf-cover-letter

## Instructions
  - Run `pip install docx2pdf` (It allows you to convert docx files to pdf.) and `pip install python-docx` (It allows you to work with Microsoft Word files)
  - Open `main.py` and change the parameters starting at line 32 to what you need it to be
  - Open the coverletter.docx file and delete the paragraphs and insert what you wish to be in the final pdf
  - Run `main.py`
  
## How It Works
  - Basically I'm just reading the word doc, copying all the data there and finding the key strings `HeaderName`, `YourEmail`, `YourNumber`, `YourWebsite`, `FooterName`and `COMPANY`and just replacing them with its respective inputs starting at line 32
  - I'm not sure if it will work for tables and images as I did not have them in my own cover letter
  - I did see it won't work if you copy and paste a horizontal line from google docs but I just did `===` + `ENTER` and it created its own horizontal line ([source](https://support.microsoft.com/en-us/office/insert-a-horizontal-line-9bf172f6-5908-4791-9bb9-2c952197b1a9#:~:text=The%20fastest%20way%20to%20add,a%20full%2Dwidth%20horizontal%20line.))   

