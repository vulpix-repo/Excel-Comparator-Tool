import pandas as pd
from fuzzywuzzy import process
import os

def perform_lookup(file_column_map):
    files = list(file_column_map.keys())
    dataframe_map = {}
    file_names = []

    #create dataframes
    for file in file_column_map:        
        dataframe_map[file] = pd.read_excel(file, usecols=file_column_map[file][1])
        dataframe_map[file].rename(columns={file_column_map[file][1][0]:'Designator'},inplace=True)
        dataframe_map[file].rename(columns={file_column_map[file][1][1]:file_column_map[file][0]},inplace=True)
        file_names.append(file_column_map[file][0])

    #join dataframes into one
    for i in range(len(dataframe_map)):
        if i==0:
            result_df = dataframe_map[files[0]]
        else:
            #use fuzzy            
            result_df = pd.merge(result_df,dataframe_map[files[i]], on='Designator', how='outer')

    #create comparison result column
    result_df['Match?'] = (result_df[file_names].eq(result_df[file_names[0]], axis=0)).all(axis=1)
    mismatches = (~result_df['Match?']).sum()

    return mismatches, result_df