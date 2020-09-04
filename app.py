from models.models import (validate_target_cells,
                           get_dir_and_workbooks,
                           get_sheet_name)
import pandas as pd, openpyxl, time, os, sys

# check if target_cells.txt file is present
# if not present will return [] and exit
target_cells = validate_target_cells()
if len(target_cells) < 1:
    print("\ntarget_cells.txt not found or it's empty, place in root dir with app.py\n")
    sys.exit()
else:
    print("\nLoaded in target cells from target_cells.txt\n")

# choose where excel files are
# returns tuple if valid, (filepath, list_of_valid_files)
dir_and_files_tuple = get_dir_and_workbooks()

# if made it this far then dir is valid and supported files found
os.chdir(dir_and_files_tuple[0])

# take in sheet name
target_sheet_name = get_sheet_name()

# iterate over all files
# if fails on any file then add error to `error_files_dict[file] = reason` and move on
error_files_dict = dict()

# change to xl dir and iterate over each wb and
# 1. try to open wb
# 2. check the sheet is present
# 3. attempt to grab the cell values
# 4. add to pandas dataframe

os.chdir(dir_and_files_tuple[0])

# soon to be columns of the dataframe
col_file_names = []
col_target_cell = []
col_target_cells_values = []

for file in dir_and_files_tuple[1]:
    try:
        # multiple try-excepts for specific error handling of opening workbooks or not present worksheets
        try:
            wb = openpyxl.load_workbook(filename=file,
                                        read_only=True,
                                        data_only=True)
        except:
            error_files_dict[file] = "Did not open. Password? Also: CSV not supported. Save as .xlsx."
            pass
        try:
            ws = wb[target_sheet_name]

            # made it this far therefore it opened the wb and found the ws
            # iterate over all the target cells in the wb and add to lists for later use in dataframe

            for cell in target_cells:
                col_file_names.append(file)
                col_target_cell.append(cell)
                try:
                    col_target_cells_values.append(ws[cell].value)
                except:
                    col_target_cells_values.append('No such cell exists')
                    error_files_dict['invalid_cell_provided'] = f"Cell '{cell}' was provided but this is not a cell"
        except:
            error_files_dict[file] = f"Could not find worksheet {target_sheet_name}"
            pass
    except:
        error_files_dict[file] = "Some unspecified error occurred"
        pass

# if made it this far then create df and output the csv
data_dict = {'File Name': col_file_names,
             'Target Cell': col_target_cell,
             'Cell Value': col_target_cells_values}

if len(data_dict['File Name']) > 1:
    df = pd.DataFrame(data=data_dict)
    # pivoting it to more user-friendly format
    df_pivot = df.pivot(index='File Name', columns='Target Cell', values='Cell Value')
    # output the csv
    df_pivot.to_csv('output.csv', index=True)
    print(f"\n###\noutput.csv exported\n###\n")
else:
    print("\nNothing output since nothing input")

if len(error_files_dict) > 0:
    print("\n!!!\nErrors on the following\n!!!\n")
    for k, v in error_files_dict.items():
        print(f"{k}: {v}")

print("\nfinished\nClosing terminal window in 10 minutes\nPress CTRL+C to exit before this.")
time.sleep(600)
