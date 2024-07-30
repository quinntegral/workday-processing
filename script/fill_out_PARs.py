import pdf2docx
from docx import Document
import os 
import re

# TODO : find way to distinguish between employee in filename and enforce that
# aggregate 3 1-month reports per employee into (now larger) PAR report without overwriting cells
# idea: change parse_workday_docx into look that takes in multiple documents and returns a larger raw_data list

def parse_workday_docx(document: Document):
    ''' helper that gets raw data from workday docx '''
    table = document.tables[1]
    data = []
    keys = None
    for i, row in enumerate(table.rows):
        text = [cell.text for cell in row.cells]
        if i == 0:
            continue
        if i == 1:
            keys = tuple(text)
        row_data = dict(zip(keys, text))
        data.append(row_data)
    return data


def fetch_workday_data(pdf_path):
    # create empty docx, convert workday output to docx, save
    docx_path = "./workday-docs/workday-data-to-parse.docx"
    open(docx_path, "w")
    document = Document()
    os.chmod(docx_path, 0o775)
    document.save(f"{docx_path}")

    pdf2docx.parse(pdf_path, docx_path)
    if not os.path.exists(docx_path):
        exit(f"Failed to create file {docx_path}.")
    document = Document(docx_path)

    # get data from document
    raw_data = parse_workday_docx(document)
    name, start_date, end_date = document.paragraphs[1].text.split('\n')
    start_date = start_date[-10:].strip()
    end_date = end_date[-10:].strip()
    name = name.strip()
    return raw_data, name, start_date, end_date


def organize_data(raw_data, name):
    ''' returns data in format [comment, date, hours logged]
    combines entries that are logged on the same day '''
    entries = []
    temp = []
    r = re.compile('[0-9]*/[0-9]*/[0-9]*')
    h = re.compile('Hours:.*')
    for index, entry in enumerate(raw_data):
        if index == 0:
            continue
        else:
            if r.match(entry['Date']) is not None:
                # date for the current entry
                if len(temp) >= 2:
                    continue
                temp.append(entry['Date'])
            elif h.match(entry['Date']) is not None:
                # hours for the current entry
                temp.append(entry['Date'][9:])
                entries.append(temp)
                # reset temp
                temp = []
            else:
                # comment for the current entry
                if len(temp) >= 2:
                    temp[0] += ', ' + entry['Comment']
                else:
                    temp.append(entry['Comment'])
    return entries


def fill_document(data, name, start_date, end_date):
    PAR_path = "./par-template/PAR-template.docx"
    PAR = Document(PAR_path)

    # fill in header info
    PAR.tables[2].rows[0].cells[1].text.replace("XX", name)
    PAR.tables[2].rows[2].cells[3].text = start_date
    PAR.tables[2].rows[2].cells[5].text = end_date

    # sublist indices: 0=description, 1=date, 2=hours
    # cells in PAR table: 0=date, 1=description, 2=hours
    for i, sublist in enumerate(data):
        if not sublist[0]:
            print(f"NOTE: {name} didn't include a description for all days worked.")
        PAR.tables[3].rows[i+1].cells[1].text = sublist[0]
        PAR.tables[3].rows[i+1].cells[0].text = sublist[1]
        PAR.tables[3].rows[i+1].cells[2].text = sublist[2]
    
    name = name.replace(" ", "")
    start_date = start_date.replace("/", "")
    filename = f'./filled-reports/PAR-{name}-{start_date}.docx'
    open(filename, "w")
    PAR.save(filename)
    os.chmod(filename, 0o666)  # gives permissions rw-rw-rw-


def main():
    for filename in os.listdir("./employee-pdfs"):
        print(f"Processing {filename}...")
        raw_data, name, start_date, end_date = fetch_workday_data(f"./employee-pdfs/{filename}")
        if raw_data:
            organized_data = organize_data(raw_data, name)
            fill_document(organized_data, name, start_date, end_date)
        else:
            exit(1, "Fetching workday data failed. Exiting")
    print("Done.")


if __name__ == "__main__":
    main()