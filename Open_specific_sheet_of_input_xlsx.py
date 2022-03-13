import os


def open_requested_sheet_of_an_excel(excel_filename, sheet_name):
    project_dir = os.path.dirname(__file__)
    excel_file_path = os.path.join(project_dir, 'input_data\\Input_data_file.xlsx')
    print("Opening {}".format(excel_filename))
    os.system("{}\\open-a-sheet.vbs {} {}".format(project_dir, excel_file_path, sheet_name))


open_requested_sheet_of_an_excel('Input_data_file.xlsx', 'Test_2')

