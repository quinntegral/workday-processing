# workday-processing
Script for transforming Workday's time-tracking export pdf to American University's Personnel Activity Report docx file.

## description
This script takes in monthly Workday time-tracking report PDFs and produces filled out PAR reports.

## how to use
To set up for use:
* From Workday, download monthly time-tracking Workday pdfs for each employee. (Any time period is fine, but monthly reduces the amount of files you have to download.)
* In the employees directory, make a folder for each employee. Then upload their Workday output pdfs to their directory.
* Ensure you have a Python virtual environment set up.
* Use the Makefile to install dependencies and run the script from the repo's root directory.
  * The "make" command will install dependencies and run the script.
  * To clean up used/generated files, run "make clean".
* Finalized PARs will be output to the filled-reports directory. Their filenames will contain the employee's name and the Workday pdf's time range.