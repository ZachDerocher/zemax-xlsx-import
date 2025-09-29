import pandas as pd
import numpy as np
# function to support reading data from excel into pandas dataframes
# organize each data block (META, SURF, ASPH, CONF, WAVE) into a dataframe, and store as dict 

def check_excel_data_key(expected_keys, keys, str1, str2):
    # check if the keys are valid
    for key in keys:
        if key not in expected_keys:
            print(f"error: unknown {str1} key '{key}'")
            print(f"excel {str2} should only contain the following values: \n{expected_keys}")
            return

def read_excel_patent_data(fn):
    df = pd.read_excel(fn, header=None)

    data_types = df[0][:].unique()

    #create a data frame dictionary to store dataframes
    lens_data = {elem : pd.DataFrame() for elem in data_types}

    for key in lens_data.keys():
        lens_data[key] = df[:][df[0] == key]

        # clean up each dataframe
        lens_data[key].columns = lens_data[key].iloc[0]
        lens_data[key] = lens_data[key][1:]  # Drop the first row as it's now the header
        lens_data[key] = lens_data[key].reset_index()
        lens_data[key] = lens_data[key].drop(lens_data[key].columns[0:2], axis=1) # drop the first dummy cols
        
        #lens_data[key] = lens_data[key].dropna(axis=1,how='all') # remove the NaN cols; removes empty 'cir' column...
        if np.nan in lens_data[key].keys():
            lens_data[key] = lens_data[key].drop(columns=[np.nan])


    # we now have a dict of dataframes, with names taken from excel column 1
    # example usage:
    #print(lens_data.keys())
    #print(lens_data['SURF'].keys())
    #print(lens_data['SURF']['r'])
    #print(lens_data['SURF']['r'][5])

    # here is an example of how we could handle errors...
    expected_lens_data_keys = ['META', 'SURF', 'ASPH', 'CONF', 'WAVE'] # excel sheet column 1 values
    expected_meta_data_keys = ['lens_unit'] # excel sheet META column headers
    expected_surf_data_keys = ['surf_num', 'r', 'd', 'nd', 'vd', 'cir'] # excel sheet SURF column headers
    expected_wave_data_keys = ['wave_num', 'wavelength_nm', 'weight'] # excel sheet WAVE column headers
    #expected_conf_operand_types = ['d_', 'fno', 'y_'] # excel sheet CONF name value types

    check_excel_data_key(expected_lens_data_keys, lens_data.keys(), "lens data", "column 1")
    check_excel_data_key(expected_meta_data_keys, lens_data['META'].keys(), "meta data", "META data column headers")
    check_excel_data_key(expected_surf_data_keys, lens_data['SURF'].keys(), "surface data", "SURF data column headers")
    check_excel_data_key(expected_wave_data_keys, lens_data['WAVE'].keys(), "wave data", "WAVE data column headers")

    
    return lens_data
