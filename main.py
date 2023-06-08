# SPL Adobe sign form tool
# Goal: Automate the adobe sign process to update a Excel or CSV file with dates and times of equipment and key

import logging
from src.file_stuff import *
from src.excel import *
from consolemenu import *
from consolemenu.items import *
from src.pdf import *



def main():
    project_root = get_project_root()
    config = get_config(project_root)
    wb_filename = config["excelfile"]
    pdf_saveasdir = config["saveas"]
    default_import = config["defaultimport"]

    # Menu
    selection_menu = SelectionMenu(["Import", "Initialize", "Test"], "Adobe Sign Tool")
    selection_menu.show()
    menu_entry_index = selection_menu.current_option
    if menu_entry_index == 0:
        # import file
        while True:
            target = input("Directory of pdf files > ")
            if os.path.isdir(target):
                break
            print("Not a directory or does not exist. Try again")
        files = [os.path.join(target,f) for f in os.listdir(target) if f.split('.')[-1]=="pdf"]
        wb = load_wb(wb_filename)
        for each in files:
            data = extract_data(each)
            data["Link to PDF"] = "=HYPERLINK(\"{}\",\"PDF Form\")".format(os.path.join(pdf_saveasdir, filename_generate(data)))
            wb = append_table(wb,data)
            save_wb(wb_filename, wb)
            copy_file(each, os.path.join(pdf_saveasdir, filename_generate(data)))

        return
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
