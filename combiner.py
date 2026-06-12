import pandas as pd

# 1. Load the Excel file containing NIC codes
# Ensure the column names match your actual data
df_nic = pd.read_excel('ALL_NIC codes.xlsx', sheet_name=0) 
df_nic.rename(columns={'Company Name': 'company_name', 'NIC codes': 'nic_code'}, inplace=True)

# 2. Load the CMIE master file
# Assuming it is pipe-delimited as per standard CMIE formatting
df_master = pd.read_csv('cpy_cin_code.dt', sep='|', engine='python', on_bad_lines='skip')

# Clean column names (convert to lowercase, strip spaces) for easier merging
df_master.columns = [col.strip().lower() for col in df_master.columns]

# Assuming df_master has columns like 'company name', 'mca cin', 'cmie company code'
# Rename them to match for merging
df_master.rename(columns={'company name': 'company_name'}, inplace=True)

# 3. Merge the datasets
# This creates a master dataframe containing Name, NIC, CIN, and CMIE Code
merged_df = pd.merge(df_nic, df_master, on='company_name', how='inner')

# 4. Filter for your target sectors (e.g., Food Agro: 10, 11, 12)
target_nics = [10, 11, 12]
food_agro_df = merged_df[merged_df['nic_code'].isin(target_nics)].copy()

print(f"Found {len(food_agro_df)} companies in the specified NIC codes.")