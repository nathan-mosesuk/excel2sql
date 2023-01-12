excel2sql

==============================

Skills:

SQL
Python
XML
xml.etree.ElementTree
Excel

==============================

Description:

A python script that converts an Excel XML spreadsheet into an sqlite database copying
the rows and columns into the database just the way they are in the Excel file. The code
automatically cleans the XML file in order for it to be readable by the python library.
It prompts the user whether to insert empty rows into the database (just as they appear in
the Excel file) or not.

Note: the code supports Excel files with 1 spreadsheet and a maximum of 20 columns.
If the user chooses to ignore empty rows, and there are empty rows present in the Excel
file, the correspondence between the rows of the Excel file and the database will be 
affected as it will ommit inserting empty rows.
The code has the potential to convert Excel XML files with multiple spreadsheets into one
unique database table or multiple tables (one for every spreadsheet for example).

==============================

Preparing the Excel file:

If the file has multiple spreadsheets besides the first one, but they are empty, they can be left 
there but if the sheets have data, it will be added to the database, with the same columns 
correspondence as in the first sheet. However, the insertion of empty rows might not be done
correctly. Data will be unaffected anayway. For optimal results, make sure the Excel file has 
only 1 spread sheet with data.

==============================

Running the code:

-From windows command prompt, execute the .py file by writing "python filename.py"

==============================

Requirements:

-Python 3.7.1 or superior installed

==============================

Output:

-Cleaned XML file (marked '-clean'), without schemas specifications
-sqlite relational database file (.sqlite) with 1 table of the number of columns specified
by the user

==============================

Files uploaded:

-python script: excel2sql.py
-Instructions on how to save xls or xlsx Excel file as XML spreadsheet
-Output files (as example): 
-1 Output example: Excel spreadsheet of 20 columns (no empty rows)
-2 Output examples: one ignoring empty rows and another inserting empty rows
   -screenshots of sqlite database
   -screenshots of windows command prompt
   -screenshots of excel xls file 
   -screenshots of xml file (before and after cleaning)


