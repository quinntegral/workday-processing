# workday-processing
Script for transforming Workday's time-tracking export pdf to American University's Personnel Activity Report docx file.

## description
This script takes in monthly Workday time-tracking report PDFs and produces filled out PAR reports.

## how to use
To set up for use:
* Upload Workday output pdfs to the employee-pdfs directory. The finalized reports will be output to the filled-reports directory.
* Follow specific naming convention TBD to distinguish between employees.
* Ensure you have a Python virtual environment set up.
* Install dependencies and run script using the Makefile (i.e. run "make" command) from the repo's root directory.