from docx import Document
from docx.shared import Pt
from pdf2docx import parse
import os 
import re

def fetch_workday_data():
    # create docx in directory
    docx_path = "./unfilled-reports/test.docx"
    document = Document()
    os.chmod(docx_path, 0o775)
    document.save(f"{docx_path}")

    # convert from pdf to docx
    pdf_path = "./employee-pdfs/jq_example.pdf" # change to loop later
    parse(pdf_path, docx_path)
    if not os.path.exists(docx_path):
        print(f"Failed to create file {docx_path}.")
        return None

    # save and parse dat
    document = Document(docx_path)
    raw_data = parse_workday_docx(document)
    organize_data(raw_data)

def organize_data(raw_data):
    # organizes data into : [comment, date, hours logged]
    # combines entries that are logged on the same day
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
                if len(temp) > 2:

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
                if len(temp) == 2:
                    print("happened")
                    temp[0] += ', ' + entry['Comment']
                temp.append(entry['Comment'])
    return entries




def parse_workday_docx(document: Document):
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
    
    for row in data:
        print("\n")
        print(row)

    return data
    # FROM ORIG: date, comment, quantity
    # just look at reported time

    # TO NEW DOC: date, detailed activity description, hours worked
    # 0 6 8




# fill the Word doc
# def fill_document(data):
#     doc_path = "/mnt/data/Personnel Activty Report Template (1).doc"
#     doc = Document(doc_path)
    
#     # find/fill in the fields (this will depend on the exact structure of your data)
#     for para in doc.paragraphs:
#         if "Name/Title:" in para.text:
#             para.text = para.text.replace("XX", data['employee_name'])
#         if "Reporting Period:" in para.text:
#             para.text = para.text.replace("(Beginning)", data['start_date'])
#             para.text = para.text.replace("(Ending)", data['end_date'])
    
#     table = doc.tables[0]  # Assuming there is only one table
#     for i, entry in enumerate(data['entries']):
#         row = table.add_row()
#         row.cells[0].text = entry['date']
#         row.cells[1].text = entry['description']
#         row.cells[2].text = str(entry['hours'])

#     # save the filled document
#     doc.save("/Filled_Personnel_Activity_Report.docx")

# main function to execute the script
def main():
    data = fetch_workday_data()
    # if data:
    #     fill_document(data)

if __name__ == "__main__":
    main()