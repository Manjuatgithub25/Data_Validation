from Utilities.excel_data_reader import data_reader
from Utilities.excel_sheet_creator import create_excel_sheet


def score_math_operations(timestamped_dir):
    table1_data = data_reader('Sheet1', "Score", 'Table1')
    table2_data = data_reader('Sheet1', "Score", "Table2")
    table_data_name = data_reader('Sheet1', "Name", "Table2")
    table_columns = ["Name", 'table1_score', 'table1_score_sum', "table1_Average_score", "table1_row_count", 'table2_score',
                     'table2_score_sum', "table2_Average_score", "table2_row_count", "Status"]
    table_data = []
    t1_sum = sum(table1_data)
    t1_avg_score = t1_sum / len(table1_data)
    t1_row_count = len(table1_data)

    t2_sum = sum(table2_data)
    t2_avg_score = t2_sum / len(table2_data)
    t2_row_count = len(table2_data)
    for i in range(len(table_data_name)):
        if i == 0:
            if t1_sum == t2_sum and t1_avg_score == t2_avg_score and t1_row_count == t2_row_count:
                row_data = [table_data_name[i], table1_data[i], t1_sum, t1_avg_score, t1_row_count, table2_data[i], t2_sum,
                            t2_avg_score, t2_row_count, "Passed"]
            else:
                row_data = [table_data_name[i], table1_data[i], t1_sum, t1_avg_score, t1_row_count, table2_data[i],
                            t2_sum,
                            t2_avg_score, t2_row_count, "Failed"]
        else:
            row_data = [table_data_name[i], table1_data[i], None, None, None, table2_data[i], None, None, None]

        table_data.append(row_data)

    create_excel_sheet(table_columns, table_data, "Score_Mathematical_operations", timestamped_dir)


def salary_math_operations(timestamped_dir):
    table1_data = data_reader('Sheet1', "Salary", 'Table1')
    table2_data = data_reader('Sheet1', "Salary", "Table2")
    table_data_name = data_reader('Sheet1', "Name", "Table2")
    table_columns = ["Name", 'table1_salary', 'table1_salary_sum', "table1_Average_salary", "table1_row_count", 'table2_salary',
                     'table2_salary_sum', "table2_Average_salary", "table2_row_count", "Status"]
    table_data = []
    t1_sum = sum(table1_data)
    t1_avg_salary = t1_sum / len(table1_data)
    t1_row_count = len(table1_data)

    t2_sum = sum(table2_data)
    t2_avg_salary = t2_sum / len(table2_data)
    t2_row_count = len(table2_data)
    for i in range(len(table1_data)):
        if i == 0:
            if t1_sum == t2_sum and t1_avg_salary == t2_avg_salary and t1_row_count == t2_row_count:
                row_data = [table_data_name[i], table1_data[i], t1_sum, t1_avg_salary, t1_row_count, table2_data[i], t2_sum,
                            t2_avg_salary, t2_row_count, "Passed"]
            else:
                row_data = [table_data_name[i], table1_data[i], t1_sum, t1_avg_salary, t1_row_count, table2_data[i],
                            t2_sum,
                            t2_avg_salary, t2_row_count, "Failed"]
        else:
            row_data = [table_data_name[i], table1_data[i], None, None, None, table2_data[i], None, None, None]

        table_data.append(row_data)

    create_excel_sheet(table_columns, table_data, "Salary_Mathematical_operations", timestamped_dir)

