from Tasks.HLD import tasks_hld
from Tasks.Mathemetical_operations import score_math_operations, salary_math_operations
from Tasks.Mismatched_data import missing_data, mismatched_data
from Tasks.column_comparison import compare_score, compare_salary
from Tasks.percentage_change import salary_change_percentage
from Utilities.excel_sheet_creator import create_timestamped_directory


def run_tasks():
    timestamped_dir = create_timestamped_directory()
    compare_score(timestamped_dir)
    compare_salary(timestamped_dir)
    score_math_operations(timestamped_dir)
    salary_math_operations(timestamped_dir)
    salary_change_percentage(timestamped_dir)
    missing_data(timestamped_dir)
    mismatched_data(timestamped_dir)
    tasks_hld(timestamped_dir)


run_tasks()
