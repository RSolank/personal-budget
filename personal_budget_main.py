import sys, os, warnings, getopt
import datetime, string, random
import openpyxl
from tabulate import tabulate


def extract_data_from_source(dir_path):
    """
    Extracts transaction data from files in source directory and returns it as dictionaries.
    """
    
    transactions_all, monthly_summary  = {}, {}
    
    # Search for Excel files in source directory.
    try:
        excel_file_exist = False
        file_list = os.listdir(dir_path)
        for file_name in file_list:
            if file_name.endswith(".xls") or file_name.endswith(".xlsx"):
                excel_file_exist = True
                file_url = dir_path + "/" + file_name

                # Importing Excel data.
                warnings.filterwarnings("ignore", ".*DrawingML*.")
                try:
                    wb = openpyxl.load_workbook(file_url)
                    ws = wb.active
                except Exception as e:
                    print(f"Error: {e}")
                    sys.exit(1)
                
                for row in ws.iter_rows(min_row=2, values_only=True):  # Assuming headers are in the first row
                    txn_date, txn_detail = row[0], row[2]
                    debit = row[5] if row[5] is not None else 0
                    credit = row[6] if row[6] is not None else 0
                    if isinstance(txn_date, (datetime.datetime, datetime.date)) and \
                       (isinstance(debit, (int, float)) or isinstance(credit, (int, float))):
                        
                        txn_key = "".join(random.choices(string.ascii_uppercase + string.digits, k=6)) # Unique key to identify individual transaction.
                        transactions_all[txn_key] = {   "txn_date" : txn_date.strftime("%d-%b-%Y"),
                                                        "txn_detail" : txn_detail,
                                                        "debit" : debit,
                                                        "credit" : credit
                                                    }
                        
                        # Consolidating transactions month-wise.
                        month_key = txn_date.strftime("%m-%Y")
                        if month_key not in monthly_summary:
                            monthly_summary[month_key] = {  "month_name" : txn_date.strftime("%B, %Y"),
                                                            "txn_keys" : [txn_key],
                                                            "debit" : debit,
                                                            "credit" : credit,
                                                            "surplus_deficit" : credit - debit
                                                        }
                        else:
                            monthly_summary[month_key]["txn_keys"].append(txn_key)
                            monthly_summary[month_key]["debit"] += debit
                            monthly_summary[month_key]["credit"] += credit
                            monthly_summary[month_key]["surplus_deficit"] += credit - debit
        
        if not excel_file_exist:
            sys.exit("No Excel file in source directory")
    
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
    
    return transactions_all, monthly_summary


def generate_report(transactions_all, monthly_summary, report_meta, sort_code):
    """
    Generates output data
    """
    final_report = []
    
    # Generates month-wise summary.
    if report_meta.get("type", "M") == "M":
        table = [[month_detail["month_name"], month_detail["debit"], \
           month_detail["credit"], month_detail["surplus_deficit"]] \
           for month_detail in monthly_summary.values()]
        if not sort_code.get("sort_req", False):
            final_report = table
        else:
            sort_by_column = 1 if sort_code.get("sort_by", "D") == "D" else 2 if sort_code.get("sort_by", "D") == "C" else 3
            descend = True if sort_code.get("order", "D") == "D" else False
            i = 0
            while i < len(table):
                if table[i][sort_by_column] == 0:
                    table.pop(i)
                else:
                    i += 1
            final_report = sorted(table, key = lambda x : x[sort_by_column], reverse = descend)
    
    # Generates a report showcasing required transactions.
    elif report_meta.get("type", "M") == "T":
        table = [[transaction["txn_date"], transaction["txn_detail"], \
           transaction["debit"], transaction["credit"]] \
           for transaction in transactions_all.values()]
        if not sort_code.get("sort_req", False):
            final_report = table
        else:
            sort_by_column = 2 if sort_code.get("sort_by", "D") == "D" else 3 if sort_code.get("sort_by", "D") == "C" else 0
            descend = True if sort_code.get("order", "D") == "D" else False
            i = 0
            while i < len(table):
                if table[i][sort_by_column] == 0:
                    table.pop(i)
                else:
                    i += 1
            final_report = sorted(table, key = lambda x : x[sort_by_column], reverse = descend)

    # Generates a report showcasing key transactions of required months.
    else:
        for month_key in monthly_summary:
            table = [[None, transactions_all[txn_key]["txn_date"], transactions_all[txn_key]["txn_detail"], \
               transactions_all[txn_key]["debit"], transactions_all[txn_key]["credit"]] \
               for txn_key in monthly_summary[month_key]["txn_keys"]]

            sort_by_column = 3 if sort_code.get("sort_by", "D") == "D" else 4 if sort_code.get("sort_by", "D") == "C" else 1
            descend = True if sort_code.get("order", "D") == "D" else False
            i = 0
            while i < len(table):
                if table[i][sort_by_column] == 0:
                    table.pop(i)
                else:
                    i += 1
            temp_report = sorted(table, key = lambda x : x[sort_by_column], reverse = descend)
            final_report.append([monthly_summary[month_key]["month_name"]])
            #print(tabulate([], headers=monthly_summary[month_key]["month_name"], tablefmt="grid"))
            final_report += temp_report[:(report_meta.get("txn_per_month", 3))]

    return final_report


def show_output(report, report_meta, view_count):
    """
    Displays the summary of transactions in a tabulated format.
    """
    if report_meta.get("type", "M") == "M":
        headers = ["Month", "Debit", "Credit", "Surplus/Deficit"]
    elif report_meta.get("type", "M") == "T":
        headers = ["Transaction_Date", "Transaction_Details", "Debit", "Credit"]
    else:
        headers = ["Month", "Transaction_Date", "Transaction_Details", "Debit", "Credit"]
        if view_count is None:
            view_count = len(report)
        else:
            cnt = 0
            for i in range(len(report)):
                if report[i][0]:
                    cnt += 1
                if cnt > view_count:
                    view_count = i
                    break

    print(tabulate(report[:view_count], headers=headers, tablefmt="grid", numalign="right", floatfmt=".2f"))


if __name__ == "__main__":
    dir_path = os.path.split(os.path.realpath(__file__))[0] + "/source_data" # Path of directory where raw data is stored.
    report_meta, sort_code, view_count = {}, {}, None

    # Short options
    short_options = "ht:s:o:c:"

    # Long options
    long_options = ["help", "report_type=", "sort_by=", "order_by=", "view_count="]

    try:
        # Parsing argument
        arguments, values = getopt.getopt(sys.argv[1:], short_options, long_options)
        
        for currentArgument, currentValue in arguments:

            if currentArgument in ("-h", "--help"):
                print ("Displaying Help")
                
            elif currentArgument in ("-t", "--report_type"):
                # Selecting output view based on user input.
                temp = currentValue.lower().split("_")
                if "month" in temp[0]:
                    if "top" in temp:
                        report_meta["type"] = "MT"
                        report_meta["txn_per_month"] = int(temp[-1]) if temp[-1].isnumeric() else 3
                        sort_code["sort_req"] = True
                    else:
                        report_meta["type"] = "M"
                elif "transaction" in temp[0]:
                    report_meta["type"] = "T"
                    view_count = 20 if view_count is None else view_count
                else:
                    report_meta["type"] = "M"
                
            elif currentArgument in ("-s", "--sort_by"):
                # Sort by column "D" : Debit, "C" : Credit, "S" : Surplus
                sort_code["sort_req"] = True
                sort_code["sort_by"] = "D" if "debit" in currentValue.lower() else "C" if "credit" in currentValue.lower() else "S"

            elif currentArgument in ("-o", "--order_by"):
                # Sort order Descending or Ascending
                sort_code["order"] = "D" if any(sub in currentValue.lower() for sub in ["descend", "decreas"]) else "A"

            elif currentArgument in ("-c", "--view_count"):
                view_count = int(currentValue) if currentValue.isnumeric() else 10
    
    except getopt.error as error:
        # Output error and return with an error code.
        print(str(error))
    
    transactions_all, monthly_summary = extract_data_from_source(dir_path)

    report = generate_report(transactions_all, monthly_summary, report_meta, sort_code)
    show_output(report, report_meta, view_count)
