# Week11
--- Repair and optimize or redo pdf downloader ---

Before running, create the folder 'pdf_download' in your working directory or change the global variable 'folder' in the file 'pdf_downloader.py' to the path you wish 10.000 pdf files to be downloaded to.

Dependencies:

openpyxl==3.0.10
PyPDF2==2.10.5
PyPDF2==3.0.1
Requests==2.31.0


When run, the file 'pdf_downloader.py' will attempt to download all pdf files listed in the excel file 'GRI_2017_2020.xlsx' from the urls listed in columns AL and AM. They will be saved in the folder 'pdf_download' in the working directory.
An excel file is also created, 'MetaData.xlsx', containing BR#'s, urls, and information about whether the pdf was downloaded or not or if the links are defective.
On consecutive runs of the file, pdf's that are already contained in the 'pdf_download' folder or has had their urls marked as defective will be skipped by the downloader.

- - - - - - - - 
'pdf_downloader.py' contains 5 functions:
 - get_rows_excel
 - save_pdf_url
 - metadata_excel
 - GRI_pdf_downloader
 - GRI_pdf_multi_downloader

get_rows_excel is a yield function, it takes an excel file and returns a generator object that yields single rows from the excel file using openpyxl.

save_pdf_url takes an url and a desired name and path. Using urllib, a request is made and headers assigned to increase odds of the download succeding. A download is attempted and the file is saved with the assigned name at the assigned path. The function has a timeout parameter for the request that is initially set at 5s.

metadata_excel checks whether the file 'MetaData.xlsx' exists in the working directory and then, using openpyxl, either creates it or loads it accordingly. Columns are named and the openpyxl workbook is returned.

GRI_pdf_downloader is an earlier single thread downloader and NO LONGER IN USE.

GRI_pdf_multi_downloader is the main downloading function and utilizes the above helper functions to loop through 'GRI_2017_2020.xlsx' and attempt to download all listed pdf files. It will attempt the first link and then if needed the second. If both fail and the errors are of a type that indicates a broken link, the pdf is marked as having defective download links in the 'MetaData.xlsx' file. While looping through rows it will skip pdfs either already present in the 'pdf_download' folder or marked as having defective download links. The function utilises multithreading, initialising the above-mentioned try-except behavior for a single row several times in parallel using the concurrent.futures module. The number of threads is set to 16, but depending on the computer this script is run on this can be set higher for faster runtime.

When the script has finished running, the runtime in seconds will be printed.















