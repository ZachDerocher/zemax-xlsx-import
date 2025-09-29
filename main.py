# read patent data from excel, and import into Zemax
# relies on a specific excel format
# uses pandas for data import, ZOSAPI for export to zemax lens file
#
# Ansys Inc.
# Zach Derocher (zach.derocher@ansys.com)
# 2025-Sept

import read_excel_patent_data
import initialize_zemax_connection
import write_patent_data_to_zemax
import os

# === USER INPUTS === #
# please customize the input (browse window, etc.)
sample = "Sample_1"
patent_file = sample + '/' + sample + '.xlsx'
out_file = os.getcwd() + '/' + sample + '/' + sample + '.zmx'
# =================== #

# read the excel file into a dict of dataframes
lens_data = read_excel_patent_data.read_excel_patent_data(patent_file)

# initialize the zemax connection (requires valid Zemax license)
zos = initialize_zemax_connection.ZosapiApplication()

# load the data into a zemax optical system and save
write_patent_data_to_zemax.write_patent_data_to_zemax(lens_data, zos, out_file)

# clean up ZOS connection
del zos
zos = None