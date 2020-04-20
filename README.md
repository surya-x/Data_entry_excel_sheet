# Data_entry_excel_sheet
 
This script is designed to convert particular PDF content into Excel sheet.

Process :
1. The program will decode the "full_pay.pdf" present in assests folder, into readable format.
2. This will scrap the required text from the pdf into excel based on different columns.
3. Then save the excel file as "full_pay.xlsx" in assests folder.


Note :

1. Paste the pdf to convert into excel into assests folder. And make sure the name of the pdf is set to "full_pay.pdf".
2. The converted Excel will be created and will be available in the folder assests with name "full_pay.xlsx"
3. On Executing the program it will ask for a date of creation which will be present below the address & city name in PDF. Make sure to write it in desired format, 
	that is  aa.aa.aaaa
	excample 11.08.2018
	
	This is used to fix the error in case if either any one of them (name, address, city) is missing from the PDF or any one of these is taking more than one line.

	
Technical info - Modules used :
1. pdfminer.six
2. numpy
3. openpyxl
4. io
5. sys