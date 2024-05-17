import pandas as pd

# Read the input Excel file
file_path ='MO L-E 2025PY - INV.xlsx' 
# Create an ExcelFile object
xls = pd.ExcelFile(file_path)

# Dictionary to store DataFrames for each sheet
dfs = {}

# Loop through each sheet in the Excel file
for sheet_name in xls.sheet_names:
    # Read the input Excel sheet
    df = pd.read_excel(xls, sheet_name)
    sheet_number = int(sheet_name[-2:])
    
    # Your existing operations for each sheet
    start_index = df.index[df.apply(lambda row: 'Benefit Package ' in str(row), axis=1)].tolist()
    end_index = df.index[df.apply(lambda row: 'Benefit Information' in str(row), axis=1)].tolist()

    if start_index and end_index:
        extracted_rows_between = df.iloc[start_index[0] + 1:end_index[0]]
        extracted_rows_between = extracted_rows_between.drop(columns=['Unnamed: 6','Unnamed: 7','Unnamed: 8','Unnamed: 9','Unnamed: 10','Unnamed: 11'])
        extracted_rows_between.columns = ['HIOS Plan ID*(Standard Component)', 'Plan Marketing Name*', 'Plan Type*','Level of Coverage*','QHP/Non-QHP*','Comments']
        extracted_rows_between.reset_index(drop=True, inplace=True)

        extracted_rows_below = df.iloc[end_index[0] + 2:]
        extracted_rows_below = extracted_rows_below.drop(columns=['Unnamed: 1'])
        extracted_rows_below.columns = ['Benefits', 'EHB', 'Is this Benefit Covered?','Quantitative Limit on Service','Limit Quantity','Limit Unit','Exclusions','Benefit Explanation','EHB Variance Reason','Excluded from In Network MOOP','Excluded from Out of Network MOOP']
        extracted_rows_below.reset_index(drop=True, inplace=True)

        # Store DataFrames in the dictionary
        dfs[sheet_name] = {'Benefit Package': extracted_rows_between, 'Benefit Information': extracted_rows_below}
