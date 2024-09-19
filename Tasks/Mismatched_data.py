import numpy as np
import pandas as pd
from Utilities.excel_data_reader import data_reader, excel_to_dict
from Utilities.excel_sheet_creator import create_excel_sheet


def missing_data(timestamped_dir):
    table1_data_name = data_reader('Sheet1', 'Name', 'Table1')
    table2_data_name = data_reader('Sheet1', 'Name', 'Table2')
    print(len(table1_data_name))
    table_columns = ['table1_name', 'table2_name', 'missing_data', 'Status']
    table_data = []

    for i in range(len(table1_data_name)):
        if table1_data_name[i] != table2_data_name[i]:
            row_data = [table1_data_name[i], table2_data_name[i], f"{table1_data_name[i]} in row {i + 2} is not as "
                                                                  f"same as {table2_data_name[i]} in row {i + 2}", "Failed"]
        else:
            row_data = [table1_data_name[i], table2_data_name[i], f"In both tables {table1_data_name[i]} are in the "
                                                                  f"same row", "Passed"]

        table_data.append(row_data)

    create_excel_sheet(table_columns, table_data, 'Missing_data_results', timestamped_dir)


def mismatched_data(timestamped_dir):
    table1_data = excel_to_dict('Sheet1', 'Table1')
    table2_data = excel_to_dict('Sheet1', 'Table2')

    df1 = pd.DataFrame(table1_data)
    df2 = pd.DataFrame(table2_data)

    attributes = ['Score', 'Salary', 'Age']
    table_columns = ['Name', 'mismatch_result', "Status"]
    table_data = []

    try:
        table1_data_name = data_reader('Sheet1', 'Name', 'Table1')
        table2_data_name = data_reader('Sheet1', 'Name', 'Table2')
        for name in range(len(table1_data_name)):
            table1_values = []
            table2_values = []

            table1_row = df1.loc[df1['Name'] == table1_data_name[name]]
            table2_row = df2.loc[df1['Name'] == table2_data_name[name]]

            if table1_row.empty or table2_row.empty:
                return ["Name not found"] * len(attributes)

            for attribute in attributes:
                if attribute in table1_row.columns:
                    value = table1_row[attribute].values[0]
                    if isinstance(value, np.int64):
                        value = int(value)
                    table1_values.append(value)
                else:
                    print("Attribute not found")

                if attribute in table2_row.columns:
                    value = table2_row[attribute].values[0]
                    if isinstance(value, np.int64):
                        value = int(value)
                    table2_values.append(value)
                else:
                    print("Attribute not found")

            # Check if values match
            if table1_values == table2_values:
                row_data = [table1_data_name[name],
                            f"{table1_data_name[name]} table1 salary, age, score are matching with {table2_data_name[name]} table2 salary, age, score", "Passed"]
                table_data.append(row_data)

            # Check if values do not match
            else:
                headers_not_equal = []
                for i in range(len(table1_values)):
                    if table1_values[i] != table2_values[i]:
                        if i == 0:
                            headers_not_equal.append("Score")
                        if i == 1:
                            headers_not_equal.append("Salary")
                        if i == 2:
                            headers_not_equal.append("Age")

                not_matched = ', '.join(headers_not_equal)
                row_data = [table1_data_name[name],
                            f"{table1_data_name[name]} table1 {not_matched} is not matching with {table2_data_name[name]} table2 {not_matched}, others are matched", "Failed"]
                table_data.append(row_data)

    except Exception as e:
        return [str(e)] * len(attributes)

    create_excel_sheet(table_columns, table_data, 'Mismatched_data_results', timestamped_dir)
