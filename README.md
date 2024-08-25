Transaction Report Generator

Overview

This Python script processes transaction data from Excel files and generates various types of reports. It can handle transactions in multiple formats, sort and filter data, and output reports in a tabulated format.

Features
    • Extracts transaction data from Excel files
    • Processes and consolidates transactions month-wise
    • Generates reports based on user-defined criteria
    • Provides options to sort and filter data
    • Outputs reports in a readable table format

Prerequisites
    • Python 3.6 or higher
    • Required Python packages:
        ◦ openpyxl
        ◦ tabulate

You can install the required packages using pip:
	pip install openpyxl tabulate

Usage

Directory Structure

Ensure your project directory has the following structure:
project/

│
├── source_data/
│   └── your_excel_files.xlsx
│
└── personal_budget_main.py

Place your Excel files in the source_data directory.

Command-Line Arguments

Run the script from the command line with the following options:
	python personal_budget_main.py [options]

Options:
    • -h, --help: Displays help information about the available options.
    • -t, --report_type: Specifies the type of report to generate. Options are:
        ◦ month for monthly summaries
        ◦ transaction for detailed transaction reports
        ◦ month_top for monthly top transactions (e.g., month_top_5)
    • -s, --sort_by: Defines the column to sort by. Options are:
        ◦ debit for sorting by debit amount
        ◦ credit for sorting by credit amount
        ◦ surplus_deficit for sorting by surplus/deficit
    • -o, --order_by: Defines the sort order. Options are:
        ◦ ascend for ascending order
        ◦ descend for descending order
    • -c, --view_count: Specifies the number of rows to display. Default is 20.

Example

To generate a monthly summary report, sorted by debit amount in descending order, and display the top 10 results:
	python personal_budget_main.py -t month -s debit -o descend -c 10

Functions

extract_raw_data_from_excel(file_url)
Extracts transaction data from an Excel file and returns a list of transactions.

process_transactions(txn_data)
Processes transaction data into a dictionary with unique transaction IDs.

consolidate_transactions_month_wise(transactions)
Consolidates transaction data month-wise and calculates monthly totals.

extract_data_from_source(dir_path)
Extracts transaction data from all Excel files in the specified directory and 
processes them.

generate_report_meta()
Generates report metadata based on command-line arguments.

generate_report(transactions, monthly_summary, report_meta)
Generates a report based on the processed transactions and report metadata.

show_output(report, report_meta)
Displays the final report in a tabulated format.

Troubleshooting
    • No Excel files found: Ensure that there are Excel files in the source_data directory and they have the correct format.
    • Errors loading workbook: Verify that the Excel files are not corrupted and are in .xlsx format.


