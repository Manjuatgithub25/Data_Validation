from Utilities.excel_data_reader import data_reader
from Utilities.excel_sheet_creator import create_excel_sheet


def compare_score(timestamped_dir):
    # Reads data based on provided attribute and name
    table_data_name = data_reader('Sheet1', "Name", "Table1")
    table1_data = data_reader('Sheet1', 'Score', 'Table1')
    table2_data = data_reader('Sheet1', 'Score', "Table2")

    table_columns = ["Name", 'table1_score', 'table2_score', 'tables_score_comparison', 'Status']
    table_data = []

    for i in range(len(table_data_name)):
        if table1_data[i] == table2_data[i]:
            row_data = [table_data_name[i], table1_data[i], table2_data[i],
                        f"In both tables the score is similar for {table_data_name[i]}", "passed"]
        else:
            greater = max(table1_data[i], table2_data[i])
            smaller = min(table1_data[i], table2_data[i])
            difference = greater - smaller
            if table1_data[i] > table2_data[i]:
                row_data = [table_data_name[i], table1_data[i], table2_data[i], f"{table_data_name[i]} in table1 has "
                                                                                f"less score compare to table2 with a"
                                                                                f" difference of {difference} ", "Failed"]
            else:
                row_data = [table_data_name[i], table1_data[i], table2_data[i], f"{table_data_name[i]} in table2 has "
                                                                                f"less score compare to table1 with a"
                                                                                f" difference of {difference} ", "Failed"]

        table_data.append(row_data)

    create_excel_sheet(table_columns, table_data, "score_comparison_results", timestamped_dir)


def compare_salary(timestamped_dir):
    table_data_name = data_reader('Sheet1', "Name", "Table1")
    table1_data = data_reader('Sheet1', "Salary", 'Table1')
    table2_data = data_reader('Sheet1', "Salary", "Table2")
    table_columns = ["Name", 't1_salary', 't2_salary', 'salary_comparison_result', "Status"]
    table_data = []
    for i in range(len(table1_data)):
        if table1_data[i] == table2_data[i]:
            row_data = [table_data_name[i], table1_data[i], table2_data[i],
                        f"In both tables the salary is similar for {table_data_name[i]}", "Passed"]
            table_data.append(row_data)
        if table1_data[i] != table2_data[i]:
            greater = max(table1_data[i], table2_data[i])

            smaller = min(table1_data[i], table2_data[i])
            difference = greater - smaller
            if table1_data[i] > table2_data[i]:
                row_data = [table_data_name[i], table1_data[i], table2_data[i], f"{table_data_name[i]} in table1 has "
                                                                                f"less salary compare to table2 with a"
                                                                                f" difference of {difference} ", "Failed"]
            else:
                row_data = [table_data_name[i], table1_data[i], table2_data[i], f"{table_data_name[i]} in table2 has "
                                                                                f"less salary compare to table1 with a"
                                                                                f" difference of {difference} ", "Failed"]
            table_data.append(row_data)

    create_excel_sheet(table_columns, table_data, "salary_comparison_results", timestamped_dir)
