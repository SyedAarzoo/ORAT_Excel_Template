import pandas as pd

# read the Excel file into a pandas dataframe
df = pd.read_excel('19AJ100429-GC3-FM-024-A1-TRENCHER-DATA-SL_T1500_Dive_025_2nd_Pass.xlsx', sheet_name='T1500', engine='openpyxl')

# get the indices of the blank rows
blank_row_indices = df.index[df.isnull().all(axis=1)]
print(blank_row_indices[-1])

# loop through the indices and create new dataframes for each section
for i in range(len(blank_row_indices)):
    if i == 0:
        section_df = df.loc[:blank_row_indices[i]]
    else:
        section_df = df.loc[blank_row_indices[i-1]+1:blank_row_indices[i]]
    
    # create a new Excel file for the section
    section_df.to_excel(f'Section{i+1}.xlsx', index=False)

section_df = df.loc[blank_row_indices[-1]+1:]
    
# create a new Excel file for the section
section_df.to_excel(f'Section{len(blank_row_indices)+1}.xlsx', index=False)
