## Dated:- 23/04/2020
## Warning: In all commands below the inverted commas are excluded ##
## Read thoroughly before proceeding.
## Last updated :- 17/02/2020
......................................................................................................

## Make sure Python is installed and the following libraries PyPDF2 and XlsxWriter are installed.
Download Links for Python:
For Windows OS:- https://www.python.org/downloads/windows/
For Mac OS:- https://www.python.org/downloads/mac-osx/
For Other Platforms:- https://www.python.org/download/other/ 

......................................................................................................

Open the folder location of Python(Not the shortcut)
open Scripts.
1. While viewing the files in the folder, hold Left Shift and right click the mouse.
2. Click on "Open Powershell Window here".
3. Type "python pip.py" to install PIP.
4. Now close the powershell window.

......................................................................................................

Open CMD 
1. Type "python -m pip install PyPDF2"
2. Type "python -m pip install XlsxWriter"

......................................................................................................
To generate the BILL Summary:

1. Make a temporary folder.
2. Copy the bill.py into the folder.
3. Copy the PDF you want to run the script on in this temporary folder.
4. Rename it as BILL.pdf.
5. While viewing the files in the folder, hold Left Shift and right click the mouse.
6. Click on "Open Powershell Window here".
7. Type "python bill.py MMM-YYYY"
   Eg- Here it is python bill.py MAY-2020
   Enter the current month and year as asked in the input format.
	Eg- MMM-YYYY
	Here - MAY-2020
	to generate heading of "BILL SUMMARY OF MAY-2020".
9. BILL.xlsx will be generated in the temporary folder.

......................................................................................................

