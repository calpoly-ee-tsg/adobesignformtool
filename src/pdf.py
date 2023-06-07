import PyPDF2 as pypdf
from excel import dataframe

def extract_data(path, form_type="checkout-agreement"):
    pdfobject = open(path, 'rb')
    pdf = pypdf.PdfReader(pdfobject)
    text = []
    for page in pdf.pages:
        text += page.extract_text().split('\n')
    result = dataframe()
    result["Name"] = text[69]
    result["Email"] = text[70]
    result["EmplID"] = text[71]
    result["Phone"] = text[72]
    result["Checkout Start"] = text[73]
    result["Checkout End"] = text[74]
    result["Advisor Name"] = text[38]
    result["Advisor Email"] = text[40]
    result["Reason"] = text[1234567]
    result["Equipment"] = text[1234567]
    result["Equipment SN"] = text[1234567]


extract_data("C:\\Users\\ee-student-lab\\Downloads\\EE Equipment Loan Request (40).pdf")
