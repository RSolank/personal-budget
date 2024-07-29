import sys
import datetime
import openpyxl
import warnings
import tabulate

def extract_xl(file_url="Account_Statement_Axis_01-04-2021_31-03-2022.xlsx"):
	warnings.filterwarnings("ignore", ".*DrawingML*.")
	wb = openpyxl.load_workbook(file_url)
	ws = wb.active
	#txn_date = [ws.cell(row=i,column=1).value for i in range(5,263)]
	#debit = [ws.cell(row=i,column=6).value for i in range(5,263)]
	#credit = [ws.cell(row=i,column=7).value for i in range(5,263)]
	txns = {}
	for row in ws.rows:
		#print(row[0].value)
		if isinstance(row[0].value,(datetime.datetime, datetime.date)) and (isinstance(row[5].value,(int, float)) or isinstance(row[6].value,(int, float))):
			#print(row[0].value,row[5].value,row[6].value)
			txns["txn_date"] = txns.get("txn_date",[]) + [row[0].value]
			txns["debit"] = txns.get("debit",[]) + [row[5].value]
			txns["credit"] = txns.get("credit",[]) + [row[6].value]
	#txns = { "txn_date" : txn_date, "debit" : debit, "credit" : credit }
	return txns


def month_wise(txns):
	monthly = {}
	for i in range(len(txns["txn_date"])):
		deb = txns["debit"][i] if txns["debit"][i] else 0
		cred = txns["credit"][i] if txns["credit"][i] else 0
		mnth = txns["txn_date"][i].strftime("%B, %Y")
		if mnth not in monthly:
		        monthly[mnth] = [deb, cred, cred-deb]
		else:
		        monthly[mnth][0] += deb
		        monthly[mnth][1] += cred
		        monthly[mnth][2] += cred - deb

	return monthly

def show_output(monthly):
	hdrs = ["Month", "Debit", "Credit", "Surplus/Deficit"]
	table = []
	for key in monthly.keys():
		table.append([key, monthly[key][0], monthly[key][1], monthly[key][2]])
	print(tabulate.tabulate(table, headers=hdrs, tablefmt="grid"))

	return

if __name__ == "__main__":
	txns = extract_xl(sys.argv[1]) if len(sys.argv) > 1 else extract_xl()
	monthly = month_wise(txns)
	show_output(monthly)