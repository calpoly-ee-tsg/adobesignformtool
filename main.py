# SPL Adobe sign form tool
# Goal: Automate the adobe sign process to update a Excel or CSV file with dates and times of equipment and key

from consolemenu import *

from src.adobesign import *
from src.excel import *
from src.file_stuff import *

# Form types (same as excel sheet names)
FORM_TYPES = ["Equipment Loan", "Controlled Lab Access", "Keyed Lab Access"]
CONFIG_FILE_NAME = "config.json"



def main():

    # get working directory
    project_root = get_project_root()

    # get files such as default import directory and other stuff from config file
    config = get_config(project_root, CONFIG_FILE_NAME)
    wb_filename = config["excelfile"]
    pdf_saveasdir = config["saveas"]
    default_import = config["defaultimport"]


    # Menu
    selection_menu = SelectionMenu(FORM_TYPES, title="Adobe Sign Import Tool", subtitle="Pick the form you wish to "
                                                                                        "import. Please do not mix "
                                                                                        "CSV files from different "
                                                                                        "kinds of forms.")
    selection_menu.show()
    # Pick the user's entry. If it is "quit", make the following equal to None.
    menu_entry = (FORM_TYPES[selection_menu.current_option] if (selection_menu.current_option < (len(FORM_TYPES) - 1)) else None)

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
            data = parse_csv(each)
            if data is not None:
                # got a good result
                logging.info("Imported {} form for {} ({})".format(menu_entry, data["Name"], data["Email"]))
                delete_file(each)
                success += 1
                wb = append_table(wb=wb, worksheet=menu_entry, data=data)
                save_wb(wb_filename, wb)
            else:
                logging.warning("Unsuccessful import of file {}".format(each))

        if success == len(files):
            print("Processed {} {} file{}.".format(len(files), menu_entry, ("" if len(files) == 1 else "s")))
        else:
            print("Processed {}/{} {} files.".format(success, menu_entry, len(files)))
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
