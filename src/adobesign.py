import logging

from src.excel import dataframe
import csv
def parse_csv(path):
    result = dataframe()
    try:
        with open(path) as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if row["SponsorName"] != "":
                    # Professor
                    result["Advisor Name"] = row["SponsorName"]
                    result["Advisor Email"] = row["email"]
                    result["Reason"] = row["Custom Field 3"]
                    result["Equipment"] = row["Custom Field 4"]
                if row["RecipientName"] != "":
                    # Student
                    result["Name"] = row["RecipientName"]
                    result["Email"] = row["email"]
                    result["EmplID"] = row["Custom Field 1"]
                    result["Phone"] = row["Custom Field 2"]
                    result["Checkout Start"] = row["Custom Field 6"]
                    result["Checkout End"] = row["Custom Field 7"]
                if row["Custom Field 8"] != "":
                    result["Equipment SN"] = row["Custom Field 8"]
        return result
    except KeyError as e:
        logging.error("Malformed CSV file...error {}".format(e))
        return None

if __name__ == "__main__":
    print("Adobe Sign CSV test")
    print(parse_csv("C:\\Users\\ee-student-lab\\Downloads\\EE Equipment Loan Request - Form Data (1).csv"))