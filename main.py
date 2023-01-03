import os.path
import pandas as pd
import qb_connect as conn
import numpy as np
import calendar
from datetime import datetime

# Output for dataframe
desired_width = 320
pd.set_option('display.width', desired_width)
np.set_printoptions(linewidth=desired_width)
pd.set_option('display.max_columns', 10)


# https://doc.qodbc.com/qodbc/ca/index.php
# pyodbc.pooling = False

def company_info():
    '''
    Pull the company name from database, it will be added to the tables exported.
    :return: company name
    '''
    cn = conn.qb_connection('QuickBooks Data')  # pyodbc.connect('DSN=QuickBooks Data',autocommit=True)
    cursor = cn.cursor()
    # Chart of accounts
    cursor.execute("SELECT CompanyName FROM Company")
    columns = [column[0] for column in cursor.description]
    result = cursor.fetchone()
    data = []
    # append data to list.
    for row in result:
        data.append(row)
    comp_info = dict(zip(columns, data))  # create a dictionary of columns to data.

    cursor.close()
    cn.close()
    return comp_info["CompanyName"]


def retrieve_output_trial_balance_by_period(year: int, month: int) -> None:
    '''
    method to create a data range base on a year and month input.
    input the year and month and create the first date and from that
    create the last day of the month. The result is a start and end date
    :param year:
    :param month:
    :return:
    '''
    start_date = datetime(year, month, 1)
    res = calendar.monthrange(year, month)
    last_date = datetime(year, month, res[1])
    # call method for run the trial balance based on the date range provided.
    run_trial_balance_proc(start_date.date(), last_date.date())


def trial_balance_year_period(year: int = 2021, to: int = 12):
    '''
    Set the year and the month to that the trial balance needs to be
    ran for so that the start and end dates can be determined.
    :param year:
    :param to:
    :return:
    '''
    year_month = [x + 2 for x in range(to)]
    for i in year_month:
        if i == 13:
            year += 1
            i = 1
            retrieve_output_trial_balance_by_period(year, i)
        else:
            retrieve_output_trial_balance_by_period(year, i)


def run_trial_balance_proc(start_date, end_date):
    '''
    exexcute the trial balance stored procedure and append data to a csv. Future versions should be able
    to upload data to a data.
    :param start_date:
    :param end_date:
    :return:
    '''

    sql = "sp_report TrialBalance show Debit_Title, Credit_Title, Label, Debit, Credit parameters DateFrom ={0}, " \
          "DateTo = {1}  , ReportBasis = 'Accrual'".format("{d'" + str(start_date) + "'}", "{d'" + str(end_date) + "'}")
    cn = conn.qb_connection('QuickBooks Data')  # pyodbc.connect('DSN=QuickBooks Data',autocommit=True)
    cursor = cn.cursor()

    cursor.execute(sql)
    columns = [column[0] for column in cursor.description]

    data = []
    for row in cursor.fetchall():
        data.append(list(row))

    cursor.close()
    cn.close()

    # create pandas dataframe
    df = pd.DataFrame(data, columns=columns)
    df['Debit_1_Title'] = df['Debit_1_Title'].replace(r'\s+|\\n', ' ', regex=True)
    df['Credit_1_Title'] = df['Credit_1_Title'].replace(r'\s+|\\n', ' ', regex=True)
    df['CompanyName'] = company_info()

    # if file does not exist write header
    if not os.path.isfile('trial_balance_all.csv'):
        df.to_csv('trial_balance_all.csv', header='columns', index=False)
    else:  # else it exists so append without writing the header
        df.to_csv('trial_balance_all.csv', mode='a', header=False, index=False)


def get_data_from_tables(company: str, table_name: str, sheet_name: str, output_file_name: str):
    '''
    #tables of interest = ('Account', 'chart_of_accounts.xlsx', 'QB_ChartOfAccounts' )
     ('Vendor', 'vendors.xlsx', "QB_Vendors")
     ('Customer', 'customers.xlsx', "QB_Customers")

     Pull the Chart of accounts and output data to a .xlsx file
    :param company:
    :return:
   '''
    cn = conn.qb_connection('QuickBooks Data')  # pyodbc.connect('DSN=QuickBooks Data',autocommit=True)
    cursor = cn.cursor()
    cursor.execute("SELECT * FROM {0}".format(table_name))
    columns = [column[0] for column in cursor.description]

    data = []
    for row in cursor.fetchall():
        data.append(list(row))

    cursor.close()
    cn.close()
    # create pandas dataframe and output data to excel
    df = pd.DataFrame(data, columns=columns)
    df['CompanyName'] = company
    df.to_excel(output_file_name, sheet_name=sheet_name, index=False)


def get_aging_detail(company: str, ap_or_ar: str = 'ap', date_macro: str = 'Today', aging_as_of: str = 'Today',
                     output_file_name: str = "ap_aging_detail.xlsx", sheet_name: str = "QB_APAgingDetail"):
    '''
    AR or AP Aging Details reports.
    :param company:
    :param ap_or_ar:
    :param date_macro:
    :param aging_as_of:
    :param output_file_name:
    :param sheet_name:
    :return:
    '''
    columns: list = []
    cn = conn.qb_connection('QuickBooks Data')  # pyodbc.connect('DSN=QuickBooks Data',autocommit=True)
    cursor = cn.cursor()

    sql: str = ""
    if ap_or_ar == "ap":
        sql = "sp_report APAgingDetail show TxnType_Title, Date_Title, RefNumber_Title, Name_Title, " \
              "DueDate_Title, Aging_Title, OpenBalance_Title, Text, Blank, TxnType, Date, RefNumber, Name, " \
              "DueDate, Aging, OpenBalance parameters DateMacro = {0}, AgingAsOf = {1}".format(date_macro, aging_as_of)

        columns = [column[0] for column in cursor.description]
    if ap_or_ar == "ar":
        sql = "sp_report ARAgingDetail show TxnType_Title, Date_Title, RefNumber_Title, PONumber_Title, " \
              "Name_Title, Terms_Title, DueDate_Title, Aging_Title, OpenBalance_Title, Text, Blank, TxnType, " \
              "Date, RefNumber, PONumber, Name, Terms, DueDate, Aging, OpenBalance parameters DateMacro = " \
              "{0}, AgingAsOf = {1}".format(date_macro, aging_as_of)

        columns = ["TxnType_Title", "Date_Title", "RefNumber_Title", "PONumber_Title", "Name_Title", "Terms_Title",
                   "DueDate_Title", "Aging_Title", "OpenBalance_Title", "Text", "Blank", "TxnType", "Date", "RefNumber",
                   "PONumber", "Name", "Terms", "DueDate", "Aging", "OpenBalance"]

    cursor.execute(sql)

    data = []
    for row in cursor.fetchall():
        data.append(list(row))

    cursor.close()
    cn.close()

    df = pd.DataFrame(data, columns=columns)
    if ap_or_ar == "ar":
        df['Text'] = df['Text'].ffill(inplace=True)
    df['CompanyName'] = company
    df.to_excel(output_file_name, sheet_name=sheet_name, index=False)


def get_ar_aging_summary(company: str, ap_or_ar: str = 'ap', date_macro: str = 'Today', aging_as_of: str = 'Today',
                         output_file_name: str = "ap_aging_detail.xlsx", sheet_name: str = "QB_APAgingDetail"):
    sql: str = ""

    cn = conn.qb_connection('QuickBooks Data')  # pyodbc.connect('DSN=QuickBooks Data',autocommit=True)
    cursor = cn.cursor()

    if ap_or_ar == "ap":
        sql: str = "sp_report ARAgingSummary show Current_Title, Amount_Title, Text, Label, Current, Amount parameters " \
                   "DateMacro = {0}, AgingAsOf = {1}".format(date_macro, aging_as_of)
    if ap_or_ar == "ar":
        sql = "sp_report ARAgingSummary show Current_Title, Amount_Title, Text, Label, Current, Amount parameters " \
              "DateMacro = {0}, AgingAsOf = {1}".format(date_macro, aging_as_of)

    columns = [column[0] for column in cursor.description]

    data = []
    for row in cursor.fetchall():
        data.append(list(row))

    cursor.close()
    cn.close()

    df = pd.DataFrame(data, columns=columns)

    df['CompanyName'] = company
    df.to_excel(output_file_name, sheet_name=sheet_name, index=False)


def get_general_ledger_details(company: str, date_macro: str = 'LastYear', summarize_by: str = 'TotalOnly'):
    cn = conn.qb_connection('QuickBooks Data')  # pyodbc.connect('DSN=QuickBooks Data',autocommit=True)
    cursor = cn.cursor()
    cursor.execute("sp_report CustomTxnDetail show TxnType, Date, RefNumber, Name, Memo, AccountNumber, Account, "
                   "Class, ClearedStatus, "
                   " SplitAccount, Debit, Credit, RunningBalance parameters DateMacro = {0}, SummarizeRowsBy = {1}".format(
        date_macro, summarize_by))

    columns = [column[0] for column in cursor.description]

    data = []
    for row in cursor.fetchall():
        data.append(list(row))

    cursor.close()
    cn.close()

    df = pd.DataFrame(data, columns=columns)
    df['CompanyName'] = company
    df.to_excel('general_ledger.xlsx', sheet_name="QB_GeneralLedger", index=False)


def main():
    trial_balance_year_period()
    company = company_info()  # get company name
    get_data_from_tables(company, 'Vendor', 'Vendors', 'Vendor.xlsx')  # vendors
    get_data_from_tables(company, 'Customer', 'Customers', 'Customers.xlsx')  # customers
    get_data_from_tables(company, 'Account', 'Accounts', 'Accounts.xlsx')  # accounts
    get_aging_detail(company, 'ap', 'Today', 'Today', "ap_aging_detail.xlsx", "QB_APAgingDetail") #ap aging detail
    get_aging_detail(company, 'ar', 'Today', 'Today', "ar_aging_detail.xlsx", "QB_ARAgingDetail") #ar aging detail
    get_general_ledger_details(company)


if __name__ == '__main__':
    main()
