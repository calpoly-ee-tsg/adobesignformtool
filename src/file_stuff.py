import logging
import os
import json
import pathlib
from src.excel import *


def get_config(project_root):
    # file format
    # - excelfile [workbook file name]
    configfilename = 'config.txt'
    configfile = os.path.join(project_root, configfilename)
    result = {
        "excelfile": None
    }
    if os.path.exists(configfile):
        # the config file exists
        f = open(configfile,)
        file = json.load(f)
        if "excelfile" in file:
            result["excelfile"] = file["excelfile"]
            logging.debug("Loaded config file \"%s\"." % (configfilename))
            return result
        else:
            logging.error("Malformed config file. Please delete it and rerun the program.")
            raise RuntimeError
    else:
        logging.info("No config file exists.")
        i = 0
        while True:
            if i > 4:
                raise RuntimeError("Too many bad filenames")
            file = input("Path to excel file (existing or new) > ")
            if os.path.exists(os.path.dirname(file)):
                # ok- dir exists
                result["excelfile"] = file
                f = open(configfile,'w')
                json.dump(result, f)
                logging.info("Created config file")
                if os.path.exists(file) and pathlib.Path(file).suffix in [".xlsx", ".xls"]:
                    # file already exists
                    break
                else:
                    # make a new file
                    wb = new_wb()
                    save_wb(file, wb)
                break
            print("Directory doesn't exist.")
            i += 1
    return result





def get_project_root():
    # returns root directory of the script
    result = os.getcwd()
    if os.path.isdir(os.path.join(result, 'src')):
        # in root folder
        return result
    else:
        result = os.path.dirname(result)
        return result
