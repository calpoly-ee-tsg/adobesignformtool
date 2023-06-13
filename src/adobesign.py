import logging

from src.excel import dataframe
import csv



def parse_csv(path, form_fields, form_kind):
    result = dataframe(form_fields, form_kind)
    try:
        with open(path) as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                # see which of these fields is not empty
                import_settings = form_fields[form_kind]["import-settings"]
                search_non_empty = [i for i in import_settings]
                for term in search_non_empty:
                    if row[term] != "":
                        # It's a non empty row for that option
                        for column_name in import_settings[term].keys():
                            result[column_name] = row[import_settings[term][column_name]]
                        break
                # if row["SponsorName"] != "":
                #     # Professor
                #     result["Advisor Name"] = row["SponsorName"]
                #     result["Advisor Email"] = row["email"]
                #     result["Reason"] = row["Custom Field 3"]
                #     result["Equipment"] = row["Custom Field 4"]
                # if row["RecipientName"] != "":
                #     # Student
                #     result["Name"] = row["RecipientName"]
                #     result["Email"] = row["email"]
                #     result["EmplID"] = row["Custom Field 1"]
                #     result["Phone"] = row["Custom Field 2"]
                #     result["Checkout Start"] = row["Custom Field 6"]
                #     result["Checkout End"] = row["Custom Field 7"]
                # if row["Custom Field 8"] != "":
                #     result["Equipment SN"] = row["Custom Field 8"]
        return result
    except KeyError as e:
        logging.error("Malformed CSV file...error {}".format(e))
        return None

if __name__ == "__main__":
    print("Adobe Sign CSV test")
    print(parse_csv("C:\\Users\\ee-student-lab\\Downloads\\EE Equipment Loan Request - Form Data (1).csv"))