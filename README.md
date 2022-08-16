# ***Director Search Converter Readme***
# Contents
## [1.0 - Overview](#overview)

## [2.0 - File Selection](#file-selection)

## [3.0 - Options](#options)

## [4.0 - Export](#export)

## [5.0 - .xlsx File](#file)



# <a name="overview"></a>1.0 - Overview:
This program is designed as a tool for the automatic conversion of multiple .pdf Dun & Bradstreet director reports into a consolidated list of .xlsx data using pdfminer3 as the primary conversion tool. 

It also allows for several filtering options to cater to the auditing work being performed and was trialled by BDO LLP for company use circa. January 2020, where this process was originally performed manually.

The files [found here](https://github.com/dwrlewis/Director-Search-Converter/tree/master/Test%20Files) match the template of the original .pdf extracts, but have had all references to director name, dates, companies, and reference numbers randomised.
## 1.1 - Interpreter Settings
This program was generated in Python 3.8.0 using the Pycharm IDE with the following interpreter settings:

|Chardet|5.0.0|
| - | - |
|Et-xmlfile|1.1.0|
|Numpy|1.23.0|
|Openpyxl|3.0.10|
|Pandas|1.4.3|
|Pdfminer3|2018.12.3.0|
|Pip|21.1.2|
|Pycryptodome|3.15.0|
|Python-datautil|2.8.2|
|Pytz|2022.1|
|Setuptools|57.0.0|
|Six|1.16.0|
|Sortedcontainers|2.4.0|
|Wheel|0.36.2|


# <a name="file-selection"></a>2.0 - File Selection:
Save all director reports into a folder and select its file path in the top left of the user interface. Once the file path has been selected it will automatically list the contents of the folder, flagging up any files without a .pdf extension.


# <a name="options"></a>3.0 - Options:
There are several options available to filter out any unnecessary data from the extracts, including:

- Financial Year End filtering – Selecting a date range to filter by will remove any lines of data with appointment dates after the last day of the year end period, as well as resignation dates occurring before the first day of the year end period
- Director Role Filtering – Filters out any lines where the role of the person is not either marked as ‘Director’, ‘LLP Designated Member’, or ‘LLP Member’, removing roles such as ‘Secretary’
- Status Filter – Removes any companies not currently marked as ‘Active’, so companies that have been dissolved will be removed
- Rename PDF – When .pdf files are initially exported from D & B, they will normally contain a random number string as its name. All files must be renamed to that of the relevant director for audit archiving purposes. Selecting this option automatically renames the file to the name of the director after data has been imported.


# <a name="export"></a>4.0 - Export:
When the conversion button is pressed, outputs will be printed to the list box as files are converted individually, displaying the total number of companies present in the file, as well as the number that remain after filtering has been applied. The total number of companies converted, as well as those remaining after filtering, is also displayed once all files are imported.

Any .pdf files with non-standard formatting will be marked up in red, but this should only occur due to a non-director report .pdf file being present in the directory, as the .pdf formatting of the export reports is highly consistent when transforming in pdfminer3.


## 4.1 - .pdf Renaming Note:
Note that if .pdf renaming has been selected as an option, it will disable the pdf conversion button until the import file path is reselected. This is because the import is dependent on the original filenames which will no longer be present, so the import path will need to be reselected if different filters are to be applied to the data.



# <a name="file"></a>5.0 - .xlsx File
Once the conversion has completed, the file will be saved to the same folder as the imports, named as ‘Consolidated Director Report.xlsx’. If there is a file present in this folder already it will have its contents overwritten.

