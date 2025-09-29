import pandas as pd
from openpyxl import Workbook

import os.path
import sys

from helpers import column_names, clean, invert, to_be_removed

pd.set_option('future.no_silent_downcasting', True)

def main(argv):
    file_path, save_path = parse_paths(argv)
 
    wb = get_data(file_path)

    banks = pd.DataFrame({
        "Bank": [],
        "Opening Balance": [],
        "Total Credit": [],
        "Total Debit": [],
        "Closing Balance": [],
        "Formula": [],
    })

    dfs = []
    for df in wb.values():
        # Grabs bank name and removes any rows with previously processed bank names in the ledger name column.
        bank_name = df.iloc[1:2, -1].values[0] if df.columns[-1] == 'Bank' else df.columns[-1]
        df = df[~(df[column_names.ln].isin(banks['Bank']))]

        # Grabs the Opening Balance and removes the first row.
        opening_balance_row = df[df.loc[:, column_names.ln] == 'Opening Balance'][column_names.la]
        opening_balance = opening_balance_row.values[0] if not opening_balance_row.empty else pd.NA
        if not opening_balance is pd.NA:
            df = df[~(df.loc[:, column_names.ln] == "Opening Balance")]
        else:
            opening_balance = 0

        # Removes all unecessary rows and columns.
        df = df.dropna(how='all', axis=0)
        df = df.loc[:, list(column_names)]
        df = df.reset_index(drop=True)

        # Changes the date and time formats.
        df[column_names.vd] = df[column_names.vd].dt.strftime('%d-%m-%Y')

        # Creates empty rows and sorts them.
        empty_rows = pd.DataFrame(to_be_removed, index=range(len(df)), columns=df.columns)
        df = pd.concat([df, empty_rows]).sort_index(kind='mergesort').reset_index(drop=True)

        # Appropriately fills in the empty rows.
        df.loc[:, column_names.ln] = df.loc[:, column_names.ln].replace(to_be_removed, clean(bank_name))
        df.loc[:, column_names.la] = df.loc[:, column_names.la].replace(to_be_removed, pd.NA).ffill()
        for i in range(len(df)):
            if df.loc[i, column_names.dr_cr] == to_be_removed:
                above = df.loc[i - 1, column_names.dr_cr]
                df.loc[i, column_names.dr_cr] = invert[above]

        # Generates summary for processed banks and sheets.
        total_credit = df[df.loc[:, column_names.dr_cr] == 'Cr'][column_names.la].sum()
        total_debit = df[df.loc[:, column_names.dr_cr] == 'Dr'][column_names.la].sum()
        closing_balance = opening_balance + total_credit - total_debit
        formula = f"Opening Balance ({opening_balance}) + Total Credit ({total_credit}) - Total Debit ({total_debit}) = Closing Balance ({closing_balance})"
        banks = pd.concat([banks, pd.DataFrame({
            "Bank": [bank_name],
            "Opening Balance": [opening_balance],
            "Total Credit": [total_credit],
            "Total Debit": [total_debit],
            "Closing Balance": [closing_balance],
            "Formula": [formula],
        })], ignore_index=True)

        dfs.append(df)

    # Concatenates all sheets into a single sheet.
    workbook = pd.concat([*dfs], ignore_index=True)
    workbook = workbook.dropna(how='all', axis=1)
    workbook = workbook.dropna(how='all', axis=0).reset_index(drop=True)

    # Creates empty rows.
    empty_rows = pd.DataFrame(to_be_removed, index=range(1, len(workbook), 2), columns=workbook.columns)
    workbook = pd.concat([workbook, empty_rows]).sort_index(kind='mergesort').reset_index(drop=True)

    # Cleans up filler values.
    workbook.replace(to_be_removed, pd.NA, inplace=True)

    save_data(workbook, banks, save_path)
    sys.exit(0)


def parse_paths(argv):
    argv_len = len(argv)

    if argv_len < 2 or argv_len > 3:
        if argv[0].endswith('.py'):
            print(f"Usage: python {argv[0].split('/').pop()} <input_path.xlsx> [<output_path>]")
        else:
            print(f"Usage: ./{argv[0].split('/').pop()} <input_path.xlsx> [<output_path>]")
        sys.exit(1)

    input_path = argv[1]
    if not input_path.endswith('.xlsx'):
        print(f"Error: Input file must be an Excel file with .xlsx extension: {input_path}")
        sys.exit(2)
    output_path = argv[2] if argv_len > 2 else 'output.xlsx'
    if not output_path.endswith('.xlsx'):
        output_path += '.xlsx'

    if not os.path.exists(input_path):
        print(f"Error: Could not find '{input_path}'")
        sys.exit(3)
    if not os.path.exists(output_path):
        Workbook(output_path).save(output_path)

    return input_path, output_path


def get_data(file_path):
    try:
        wb = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
    except PermissionError:
        print(f"Error: Permission denied for '{file_path}'")
        sys.exit(4)
    except Exception as err:
        print(f"An unexpected error has occured: {err}.")
        sys.exit(5)

    return wb


def save_data(workbook, banks_summary, save_path):
    try:
        with pd.ExcelWriter(save_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            workbook.to_excel(writer, index=False, sheet_name='Accounting Vouchers')
            banks_summary.to_excel(writer, index=False, sheet_name='Banks Summary')
    except PermissionError:
        print(f"Error: Permission denied for '{save_path}'. Please close the file if it is open in another application.")
        sys.exit(4)
    except Exception as err:
        print(f"An unexcpected error has occured: {err}.")
        sys.exit(5)

    print(f"Workbook saved successfully to '{save_path}'!")


if __name__ == "__main__":
    main(sys.argv)
