# read patent data from excel, and import into Zemax
# relies on a specific excel format
# uses pandas for data import, ZOSAPI for export to zemax lens file
#
# Ansys Inc.
# Zach Derocher (zach.derocher@ansys.com)
# 2025-Sept

import os
import tkinter
from tkinter.filedialog import askopenfilename

import read_excel_data
import initialize_zemax_connection
import write_data_to_zemax

# Query user for excel data
root = tkinter.Tk()
root.withdraw() #use to hide tkinter window
excel_file = askopenfilename(initialdir=os.getcwd(), filetypes=[("excel files", "*.xlsx")], title='Please select an excel file')
if (len(excel_file) < 1) or (excel_file is None):
    #print(f"You chose {excel_file}")
    print("ERROR: please choose a valid excel file")
    os._exit(0)

# set the outfile name based on the selected file
out_file = excel_file[0:-5] + '_ZemaxImport.zmx'

# read the excel file into a dict of dataframes
lens_data = read_excel_data.read_excel_patent_data(excel_file)

# initialize the zemax connection (requires valid Zemax license)
zos = initialize_zemax_connection.ZosapiApplication()

# load the data into a zemax optical system and save
write_data_to_zemax.write_patent_data_to_zemax(lens_data, zos, out_file)

# clean up ZOS connection
del zos
zos = None