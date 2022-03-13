import os


def open_xlsx_file(excel_filename):
    project_dir = os.path.dirname(__file__)
    excel_file_path = os.path.join(project_dir, 'input_data\\Input_data_file.xlsx')
    print("Opening {}".format(excel_filename))
    os.system("start excel {}".format(excel_file_path))


open_xlsx_file('Input_data_file.xlsx')

