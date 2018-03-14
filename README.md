# NyATools
Tools to work around limitations in NyA. NyA is the web based system used in Sweden to manage the admission to university degrees. If you use NyA to rank applications, you will probably need to export some of the information and work locally. Here are some tools that may simplify that process.

## CompareNyAExports.py
Motivation: the admission office updates the information in NyA continuously during the admission period. If you work with exported data, you will need a method to update your local data on a regular basis.

This is a python script that can be run with:

`python CompareNyAExports.py oldExport.xlsx newExport.xlsx diff.xlsx`

and compares two exel files exported from NyA at different time steps and produces the following output in the terminal:
1) a list of applications that are in `oldExport.xlsx` but not in `newExport.xlsx`
2) a list of applications that are in `newExport.xlsx` but not in `oldExport.xlsx`
3) the applications that are in both files but with different personal numbers
The script also saves the rows that should be added to your local file in `diff.xlsx`. It will also parse the last column that contains comma separated fields with information about the degree, total number of credits, country and university. In this last case, some parsing is performed to avoid cases when commas are used within the same field.

NOTE: NyA exports in `xls` format, whereas the script only works with `xlsx`. When you export from NyA, you will have to open the file in Excel or LibreOffice and save as "Microsoft Excel 2007-2013 XML (.xlsx)" format.

### This is a typical working flow:
The first time
1) export the data from NyA (it will download a file called `excel` without extention)
2) open the file in your favourite Excel variant and save it as "Microsoft Excel 2007-2013 XML (.xlsx)" format. It is good practice to include the date and time in the file name because it might come handy later, for example `export2018-03-14_2040.xlsx`
3) Remove all the data from the file leaving only the first header row. Save this file as `empty.xlsx`
4) run `python CompareNyAExports.py empty.xlsx export2018-03-14_2040.xlsx diff2018-03-14_2040.xlsx`
You can now copy the rows in `diff2018-03-14_2040.xlsx` into your working document.

The following times:
1) export a new version, save to `xlsx` format, for example to `export2018-03-15_1030.xlsx`
2) run `python CompareNyAExports.py export2018-03-14_2040.xlsx export2018-03-15_1030.xlsx diff2018-03-15_1030.xlsx`
Again, you can now add the rows in `diff2018-03-15_1030.xlsx` into your working document, and fix the personal numbers that changed in the meantime by hand.
