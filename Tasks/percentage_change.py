from Utilities.excel_data_reader import data_reader
from Utilities.excel_sheet_creator import create_excel_sheet


def salary_change_percentage(timestamped_dir):
    table_data_name = data_reader('Sheet1', "Name", "Table1")
    table1_data = data_reader('Sheet1', "Salary", 'Table1')
    table2_data = data_reader('Sheet1', "Salary", "Table2")
    table_columns = ['Name', 'table1_salary', 'table2_salary', 'salary_difference_percentage', "Status"]
    table_data = []

    for i in range(len(table1_data)):
        change_in_salary_percentage = ((table2_data[i]-table1_data[i])/table1_data[i])*100

        if table2_data[i] > table1_data[i]:
            result = f"table2 salary is {change_in_salary_percentage}% more than table1 salary"
        else:
            result = f"table1 salary is {change_in_salary_percentage}% more than table2 salary"

        if change_in_salary_percentage == 0:
            row_data = [table_data_name[i], table1_data[i], table2_data[i], result, "Passed"]
        else:
            row_data = [table_data_name[i], table1_data[i], table2_data[i], result, "Failed"]

        table_data.append(row_data)

    create_excel_sheet(table_columns, table_data, "percentage_change", timestamped_dir)
