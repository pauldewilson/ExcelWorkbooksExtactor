"""
models.py will contain all models associated with this app
"""
import os


def validate_target_cells(tgt_cells=None):
    # try to open the target_cells txtfile and if file doesn't exist, flag error
    # not checking if the cells are valid else it would require 17bn combinations
    try:
        with open('target_cells.txt', 'r') as txtfile:
            tgt_cells = [row.replace('\n', '') for row in txtfile.readlines()]
            return tgt_cells
    except:
        print("Target cells not found.")
        tgt_cells = []
        return tgt_cells


def get_dir_and_workbooks(filepath=None):
    """
    Recursive function
    Asks for user input until valid dir path is provided
    Also ensures the dir path has at least one supported Excel format present
    """

    list_supported_extensions = ['.xlsx', '.xlsm', '.xltx', '.xltm']

    def check_for_xl_files(dirpath):
        """
        Returns False if directory contains no supported extensions
        Else returns tuple, (filepath, supported_files_found)
        """
        valid_xl_files = [f for f in os.listdir(dirpath) if os.path.splitext(f)[1].lower() in list_supported_extensions]
        if len(valid_xl_files) < 1:
            return False
        else:
            return valid_xl_files

    if filepath is None:
        filepath = str(input(r'Paste directory of Excel workbooks: '))
        return get_dir_and_workbooks(filepath=filepath)
    elif filepath is not None:
        if os.path.isdir(filepath) is True:
            supported_files_found = check_for_xl_files(dirpath=filepath)
            if supported_files_found is False:
                print(f"No supported Excel files found, script supports: {list_supported_extensions}")
                return get_dir_and_workbooks(filepath=None)
            else:
                return filepath, supported_files_found
        else:
            # executes if no valid filepath is given in first instance
            print("Filepath not found")
            return get_dir_and_workbooks(filepath=None)


def get_sheet_name(sheet=None):
    """
    Recursively asks for sheet name until one provided
    No other validation checks performed in this version
    """
    if sheet is None:
        sheet = str(input("\nType the sheet name with the target data: "))
        return get_sheet_name(sheet=sheet)
    elif sheet is not None:
        return sheet
