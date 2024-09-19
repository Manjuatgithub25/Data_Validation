from Utilities.excel_sheet_creator import create_excel_sheet, update_status_sheet


def tasks_hld(timestamped_dir):
    table_columns = ['sheet_name', 'High_level_results', 'comments']
    update_status_sheet('Tasks_HLD', table_columns, timestamped_dir)
