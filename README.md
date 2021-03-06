# MDR-app Readme
The purpose of this app is to combine large amounts of PDF files in a logical systematic way, particluarly for MDR (Master Document Register) and Workpacks.

## First use environment setup
1. Run  `pip install pipenv`
2. Run `pipenv install`. If the debug is not picking up the pipenv, follow this: https://stackoverflow.com/questions/72115439/python-pipenv-not-display-in-the-python-interpreter

## How to use
* Create a copy of `documents\My First MDR`
* Replace the templates in the `documents\templates` folder if you have specific templates you need to use.
* Edit 'doc_properties.xlsx'. Remember to choose from the drop down what type of document you want it to be.
* Create folders with names that match the Field in TableofContens sheet. In these folders can be non-pdf documents, but you manually need to convert them to PDF.
* Run the MDR App `pipenv run python mdr-app.py` to combine it all folders into one structured document.

## Code customisations
* If you want to add a new field to the Word document, in Word click Insert -> Field -> Mergefield. In Field properties \ Field name put in the name of the field to merge in (corresponding updates need to be added to the Excel file and Python code).

## Approach to app design
Approach idea 1 was selected to start with becasue MS Word documents were more friendly for external users to work with.

### Approach idea 1
1. Planning to use docx-mailmerge as shown in tutorial https://pbpython.com/python-word-template.html to be the basis for creating the reuqired pages for the MDR.
2. Convert Word docs to PDF.
3. To work with the PDF files and join them together use PyPDF3 from this tutorial here https://automatetheboringstuff.com/chapter13/.
    3.1. Merging pdf files: https://github.com/sfneal/PyPDF3/blob/master/Sample_Code/basic_merging.py
    3.2. Page numbering: https://stackoverflow.com/questions/2739159/inserting-a-pdf-file-in-latex/2740296#2740296

### Alternative approach (not used)
1. Use LaTeX for the templates. In Python use Jinja2 (for template modification).
2. Easily converts to PDF.
3. Use PyPDF3 for working with PDFs.

## Known issues
None

## Development next steps
* Have some progess icon on screen so user knows mdr is building; currently it shows nothing
* Have an error logging system setup
* Error handling; to alert user if they do an incorrect step__
* Get pdf bookmarks to default to full screen; currently depends on user settings (SL, 29/4/20)
* 1-Date in properties has a ' before it to keep as string. Let it be a date and handle that in program
* Sub-section indexes (traceability register); auto-generate to save time
* Remove all PDF bookmarks before compiling final PDF.
* Look at new python PDF engine? Currently have error on some files "PyPDF3.utils.PdfReadError: EOF marker not found" requiring re-print as pdf to work. Look at search results: https://www.google.com/search?q=PyPDF3.utils.PdfReadError%3A+EOF+marker+not+found&rlz=1C1CHBF_enAU939AU939&oq=PyPDF3.utils.PdfReadError%3A+EOF+marker+not+found&aqs=chrome..69i57.135j0j7&sourceid=chrome&ie=UTF-8
* Re-work the logic so it first runs through and indexes all the pdf documents to be added, rather than doing multiple times.

## Completed steps
* 3/12/21 Made program generic so that it can also be for workpacks. Updated templates. Updated so non-pdf documents can sit in folders.
* Improved user interface; show steps, link to procedure
* Add bookmarks for files; allows much easier browsing of MDR
* Get template files incorporated in pyinstaller (https://pyinstaller.readthedocs.io/en/stable/spec-files.html) --> up to app not finding /templates/*
* Package as .exe for distribution
* GUI with wx (how to learn: https://wiki.wxpython.org/How%20to%20Learn%20wxPython) (testing wxFormBuilder to speed up development)
* Get basic program working.
* Investigate why .exe tested in \test_exe is not working

## Building and Running Program

PyInstaller is no longer reccomended for compiling data into a single .exe file as antivirus programs often block the .exe file running.