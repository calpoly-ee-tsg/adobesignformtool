from tika import parser
from src.excel import dataframe
import logging


def extract_data(path, form_type="checkout-agreement"):
    raw = parser.from_file(path)
    text = [item for item in raw["content"].split('\n') if item != '']
    result = dataframe()
    if form_type == 'checkout-agreement':
        name_index = text.index("Returned Date:                            Initials:")+2
        while (text[name_index] == text[name_index + 1]):
            name_index += 1
        result["Name"] = text[name_index]
        result["Email"] = text[name_index + 1]
        result["EmplID"] = text[name_index + 2]
        result["Phone"] = text[name_index + 3]
        result["Checkout Start"] = text[name_index + 4].split(' ')[0]
        result["Checkout End"] = text[name_index + 4].split(' ')[1]
        result["Advisor Name"] = text[30]
        result["Advisor Email"] = text[32]
        result["Reason"] = find_custom_field(text, 'Custom Field 3: ')
        result["Equipment"] = find_custom_field(text, 'Custom Field 4: ')
        result["Equipment SN"] = find_custom_field(text, 'Custom Field 8: ')
        logging.info("Successfully extracted data from {}".format(path))
        return result
    else:
        raise NotImplementedError


def find_custom_field(text, search_term):
    return [i for i in text if search_term in i][0].split(search_term)[1]


# extract_data("C:\\Users\\ee-student-lab\\Downloads\\EE Equipment Loan Request (40).pdf")
