import requests
from docx import Document
from docx.shared import Pt

# fetch data from Workday API
def fetch_workday_data():
    api_url = "https://api.workday.com/timeTracking"
    headers = {
        "Authorization": "Bearer YOUR_API_KEY",
        "Content-Type": "application/json"
    }

    response = requests.get(api_url, headers=headers)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Failed to retrieve data: {response.status_code}")
        return None

# fill the Word doc
def fill_document(data):
    doc_path = "/mnt/data/Personnel Activty Report Template (1).doc"
    doc = Document(doc_path)
    
    # find/fill in the fields (this will depend on the exact structure of your data)
    for para in doc.paragraphs:
        if "Name/Title:" in para.text:
            para.text = para.text.replace("XX", data['employee_name'])
        if "Reporting Period:" in para.text:
            para.text = para.text.replace("(Beginning)", data['start_date'])
            para.text = para.text.replace("(Ending)", data['end_date'])
    
    table = doc.tables[0]  # Assuming there is only one table
    for i, entry in enumerate(data['entries']):
        row = table.add_row()
        row.cells[0].text = entry['date']
        row.cells[1].text = entry['description']
        row.cells[2].text = str(entry['hours'])

    # save the filled document
    doc.save("/Filled_Personnel_Activity_Report.docx")

# main function to execute the script
def main():
    data = fetch_workday_data()
    if data:
        fill_document(data)

if __name__ == "__main__":
    main()