# SPL Adobe sign form tool
# Goal: Automate the adobe sign process to update a Excel or CSV file with dates and times of equipment and key

import logging
from src.file_stuff import *
from src.excel import *
from consolemenu import *
from consolemenu.items import *
def main():
    project_root = get_project_root()
    config = get_config(project_root)
    wb_filename = config["excelfile"]

    # Menu
    selection_menu = SelectionMenu(["Import", "Initialize", "Test"], "Adobe Sign Tool")
    selection_menu.show()
    menu_entry_index = selection_menu.current_option
    if menu_entry_index == 0:
        # import file
        raise NotImplementedError
    elif menu_entry_index == 1:
        # init
        wb = load_wb(wb_filename)
        wb = initialize_workbook(wb)
        save_wb(wb_filename, wb)
        print("Initialized workbook. Please open, repair, and save as to continue.")
        print("Press enter when done.")
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
    else:
        return
    input()
    main()




if __name__ == "__main__":
    logging.basicConfig()
    main()


