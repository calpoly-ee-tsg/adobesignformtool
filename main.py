# SPL Adobe sign form tool
# Goal: Automate the adobe sign process to update a Excel or CSV file with dates and times of equipment and key

from consolemenu import *

from src.adobesign import *
from src.excel import *
from src.file_stuff import *

# Form types (same as excel sheet names)
FORM_FIELDS = {
    "Equipment Loan" : {
        # all final fields and default values
        "dataframe" : {
            "Name": "error",
            'Email': "error",
            'EmplID': "error",
            'Phone': "error",
            'Checkout Start': "error",
            'Checkout End': "error",
            'Remaining': "=[@[Checkout End]]-TODAY()",
            'Advisor Name': "error",
            'Advisor Email': "error",
            'Reason': "error",
            'Equipment': "error",
            'Equipment SN': "error",
        },
        # When the key is non-empty, it will save the nested Value in the dataframe by looking up the nested Key
        "import-settings": {
            "SponsorName": {
                "Advisor Name": "SponsorName",
                "Advisor Email": "email",
                "Reason": "Custom Field 3",
                "Equipment": "Custom Field 4"
            },
            "RecipientName": {
                "Name": "RecipientName",
                "Email": "email",
                "EmplID": "Custom Field 1",
                "Phone": "Custom Field 2",
                "Checkout Start": "Custom Field 6",
                "Checkout End": "Custom Field 7"
            },
            "Custom Field 8": {
                "Equipment SN": "Custom Field 8"
            }
        }
    },
    "Controlled Lab Access": {
        # all final fields and default values
        "dataframe" : {
            "Students": "error",
            'Access Start': "error",
            'Access End': "error",
            'Remaining': "=[@[Access End]]-TODAY()",
            'Advisor Name': "error",
            'Advisor Email': "error",
            'Room': "error",
            'Reason': 'error'
        },
        # When the key is non-empty, it will save the nested Value in the dataframe by looking up the nested Key
        "import-settings": {
            "SponsorName": {
                "Students": "Custom Field 6",
                'Access Start': "Custom Field 3",
                'Access End': "Custom Field 4",
                'Advisor Name': "SponsorName",
                'Advisor Email': "email",
                'Room': "Custom Field 2",
                'Reason': "Custom Field 5",

            },
        }
    },
    "Keyed Lab Access": {
        # all final fields and default values
        "dataframe" : {
            "Name": "error",
            'Email': "error",
            'EmplID': "error",
            'Phone': "error",
            'Key Required Date': "error",
            'Key Return Date': "error",
            'Remaining': "** please extend the excel table",
            'Advisor Name': "error",
            'Advisor Email': "error",
            'Room': "error",
            'Class': 'error',
            'Reason': "error",
            'Other Students': "error",
        },
        # When the key is non-empty, it will save the nested Value in the dataframe by looking up the nested Key
        "import-settings": {
            "SponsorName": {
                "Advisor Name": "SponsorName",
                "Advisor Email": "email",
                "Reason": "Custom Field 10",
                "Other Students": "Custom Field 11",
                "Room": "Custom Field 4",
                'Key Required Date': "Custom Field 6",
                'Key Return Date': "Custom Field 7",
                'Class': 'Custom Field 9',
                'Name': 'Form Field 8'
            },
            "RecipientName": {
                "Email": "email",
                "EmplID": "Custom Field 14",
                "Phone": "Custom Field 15"
            }
        }
    },
    "Storage Loan": {
        # all final fields and default values
        "dataframe" : {
            "Storage Designator": "=[@[Room]]&\"-\"&[@[Drawer]]",
            "Room": "error",
            'Drawer': "error",
            'Key Code': "error",
            'Issued Date': "error",
            'Student Name': "error",
            'Email': 'error',
            'Phone': "error",
            'EmplID': "error",
            'Project': "error",
            'Access Start': 'error',
            'Access End': "error",
            'Remaining': '**Please extend the excel table'
        },
        # When the key is non-empty, it will save the nested Value in the dataframe by looking up the nested Key
        "import-settings": {
            "SponsorName": {
                'Student Name': "SponsorName",
                'Email': 'email',
                'Phone': "Custom Field 3",
                'EmplID': "Custom Field 4",
                'Project': "Custom Field 5",
                'Access Start': 'Custom Field 6',
                'Access End': "Custom Field 7"
            },
            "initials": {
                "Room": "Custom Field 8",
                'Drawer': "Custom Field 9",
                'Key Code': "Custom Field 10",
                'Issued Date': "Custom Field 11"
            }
        }
    },
    "Keyed Lab Access (Legacy)": {
        # all final fields and default values
        "dataframe" : {
            "Name": "error",
            'Email': "error",
            'EmplID': "error",
            'Phone': "error",
            'Key Required Date': "error",
            'Key Return Date': "error",
            'Remaining': "** please extend the excel table",
            'Advisor Name': "error",
            'Advisor Email': "error",
            'Room': "error",
            'Class': 'error',
            'Reason': "error",
            'Other Students': "error",
        },
        # When the key is non-empty, it will save the nested Value in the dataframe by looking up the nested Key
        "import-settings": {
            "SponsorName": {
                "Advisor Name": "SponsorName",
                "Advisor Email": "email",
                "Reason": "Custom Field 7",
                "Room": "Custom Field 3",
                'Key Required Date': "Custom Field 5",
                'Key Return Date': "Custom Field 4",
                'Class': 'Custom Field 8',
                'Name': 'Custom Field 6'
            },
            "RecipientName": {
                "Email": "email",
                "EmplID": "Custom Field 11",
                "Phone": "Custom Field 15"
            }
        }
    }

}
CONFIG_FILE_NAME = "config.json"



def main():

    # get working directory
    project_root = get_project_root()

    # get files such as default import directory and other stuff from config file
    config = get_config(project_root, CONFIG_FILE_NAME)
    wb_filename = config["excelfile"]
    pdf_saveasdir = config["saveas"]
    default_import = config["defaultimport"]

    # Get form types (names)
    form_types = [i for i in FORM_FIELDS.keys()]


    # Menu
    selection_menu = SelectionMenu(form_types, title="Adobe Sign Import Tool", subtitle="Pick the form you wish to "
                                                                                        "import. Please do not mix "
                                                                                        "CSV files from different "
                                                                                        "kinds of forms.")
    selection_menu.show()
    # Pick the user's entry. If it is "quit", make the following equal to None.
    menu_entry = (form_types[selection_menu.current_option] if (selection_menu.current_option < (len(form_types))) else None)

    if menu_entry is None:
        return None
        # quit

    else:
        # Get the import directory
        target = file_prompt(text="Please enter import directory for CSV files\nPress [enter] to use default directory \'{}\'\n > ".format(default_import), must_exist=True, kind='dir', max_retries=4, default=default_import)
        # Find all importable files
        files = [os.path.join(target,f) for f in os.listdir(target) if f.split('.')[-1]=="csv"]
        # Load the workbook
        wb = load_wb(wb_filename)
        # Count number of successful imports
        success = 0
        for each in files:
            data = parse_csv(each, FORM_FIELDS, menu_entry)
            if data is not None:
                # got a good result
                try:
                    logging.info("Imported {} form under {}.".format(menu_entry, data["Email"]))
                except KeyError:
                    logging.info("Imported {} form.".format(menu_entry))
                delete_file(each)
                success += 1
                wb = append_table(wb=wb, data=data, form_fields=FORM_FIELDS, form_kind=menu_entry, worksheet=menu_entry)
                save_wb(wb_filename, wb)
            else:
                logging.warning("Unsuccessful import of file {}".format(each))

        if success == len(files):
            print("Processed {} {} file{}.".format(len(files), menu_entry, ("" if len(files) == 1 else "s")))
        else:
            print("Processed {}/{} {} files.".format(success, len(files), menu_entry))
        # open the excel file
        print("Opening file in Excel...")
        open_file_in_windows(wb_filename)
        return None


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    while True:
        returnvalue = main()
        if returnvalue is None:
            break
