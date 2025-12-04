import pandas as pd
import numpy as np
from rapidfuzz import process,fuzz 
import os

def perform_lookup(file_column_map):
    files = list(file_column_map.keys())
    dataframe_map = {}
    file_names = []

    #create dataframes
    for file in file_column_map:        
        dataframe_map[file] = pd.read_excel(file, usecols=file_column_map[file][1])

        #rename lookup column into Designator
        dataframe_map[file].rename(columns={file_column_map[file][1][0]:'Designator'},inplace=True)

        #rename compare column into file name
        file_names.append(file_column_map[file][0])
        dataframe_map[file].rename(columns={file_column_map[file][1][1]:file_column_map[file][0]},inplace=True)

    #join dataframes into one
    for i in range(len(dataframe_map)):
        if i==0:
            #move Designtor column in front
            base_df = dataframe_map[files[0]]
            des_col = base_df.pop('Designator')
            base_df.insert(0, des_col.name, des_col)
            result_df = base_df  

        else:
            # fuzzy matching
            # if none, don't replace
            # if score < 80, don't replace
            # if score > 80, match 
            current_df = dataframe_map[files[i]]

            for idx, row in current_df.iterrows():
                current_ref = row['Designator']
                if pd.isna(current_ref):
                    pass
                else:
                    extracted= process.extractOne(current_ref, base_df['Designator'], scorer = fuzz.WRatio)
                    if extracted[1] < 85:
                        pass
                    else:
                        current_df['Designator'] = current_df['Designator'].replace(current_ref,extracted[0])
                
            #current_df['Designator'] = current_df['Designator'].apply(lambda x: process.extractOne(x, base_df['Designator'], scorer = fuzz.token_set_ratio)[0]) 
            result_df = pd.merge(result_df,current_df, on='Designator', how='outer')

    #create comparison result column
    result_df['Match?'] = (result_df[file_names].eq(result_df[file_names[0]], axis=0)).all(axis=1)
    result_df.sort_values('Match?', ascending = True, inplace=True)
    mismatches = (~result_df['Match?']).sum()
    return mismatches, result_df