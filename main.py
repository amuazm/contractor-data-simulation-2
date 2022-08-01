import shutil
from openpyxl import Workbook, load_workbook
import datetime
import random
import numpy as np
from dateutil import relativedelta
import collections

# Copy dataset to refer from as backup
src = "./Files/Input/Budget Dataset Modified.xlsx"
dest = "./Files/Output/Data.xlsx"
shutil.copyfile(src, dest)

# Create new empty workbook to start from
wb_result = Workbook()
# Sheets
wb_result.active.title = "Project Details"
ws_project_details = wb_result["Project Details"]

# Project details page
if True:
    # Load dataset to refer from
    wb_data = load_workbook(dest)
    # Copy all entire Project Details sheet row by row from "data" to "result"
    for row in wb_data["Project Details"].iter_rows():
        l = []
        for cell in row:
            l.append(cell.value)
        ws_project_details.append(l)
    wb_result.close()
    # Styles
    for row in ws_project_details.iter_rows(min_row=2):
        row[11].style = "Currency"

# Save result
wb_result.save("./Files/Output/Result.xlsx")