## Check for file name inconsistencies

import pandas as pd
import os

path = 'Excelicious/Scripts/test'

def get_name_files(path):
    """ save the file names in a list """
    files = []
    for file in os.listdir(path):        
        files.append(file)
    return files

def check_date(files: list) -> list:
    """ Check for data entry inconsistencies """
    issues = []

    for file in files:
        temp_file = file.split("_")
        if len(temp_file[0]) < 6 or len(temp_file[0]) > 6:
            issues.append(file)
    
    return issues