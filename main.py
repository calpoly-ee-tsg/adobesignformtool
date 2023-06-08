# SPL Adobe sign form tool
# Goal: Automate the adobe sign process to update a Excel or CSV file with dates and times of equipment and key

from consolemenu import *

from src.adobesign import *
from src.excel import *
from src.file_stuff import *


def main():

    # get working directory
    project_root = get_project_root()

    # get files such as default import directory and other stuff from config file
    config = get_config(project_root)
    wb_filename = config["excelfile"]
    pdf_saveasdir = config["saveas"]
    default_import = config["defaultimport"]

    # Menu
    selection_menu = SelectionMenu(["Import", "Initialize", "Test"], "Adobe Sign Tool")
    selection_menu.show()
    menu_entry_index = selection_menu.current_option

    if menu_entry_index == 0:
        # import a bunch of files
        target = file_prompt(text="Please enter import directory for CSV files\nPress [enter] to use default directory \'{}\'\n > ".format(default_import), must_exist=True, kind='dir', max_retries=4, default=default_import)
        files = [os.path.join(target,f) for f in os.listdir(target) if f.split('.')[-1]=="csv"]
        wb = load_wb(wb_filename)
        success = 0
        for each in files:
            data = parse_csv(each)
            if data is not None:
                # got a good result
                delete_file(each)
                success += 1
                wb = append_table(wb, data)
                save_wb(wb_filename, wb)
            else:
                logging.info("Unsuccessful import of file {}".format(each))

        if success == len(files):
            print("Processed {} file{}.".format(len(files), ("" if len(files) == 1 else "s")))
        else:
            print("Processed {}/{} files.".format(success, len(files)))
        # open the excel file
        print("Opening file in Excel...")
        open_file_in_windows(wb_filename)
        return None
    elif menu_entry_index == 1:
        # init
        wb = load_wb(wb_filename)
        wb = initialize_workbook(wb)
        save_wb(wb_filename, wb)
        print("Initialized workbook. Please open, repair, and save as to continue.")
        print("Press enter when done.")
        return True
    elif menu_entry_index == 2:
        # test
        wb = load_wb(wb_filename)
        test = {"Name": "Alex",
                'Email': "1Alex",
                'EmplID': "2Alex",
                'Phone': "3Alex",
                'Checkout Start': "4Alex",
                'Checkout End': "5Alex",
                'Remaining': "6Alex",
                'Advisor Name': "7Alex",
                'Advisor Email': "8Alex",
                'Reason': "9Alex",
                'Equipment': "10Alex",
                'Equipment SN': "11Alex"
                }
        wb = append_table(wb, test)
        save_wb(wb_filename, wb)
        print("Added test frame to spreadsheet.")
        return True
    else:
        return None


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    while True:
        returnvalue = main()
        if returnvalue is None:
            break
