import pandas as pd
# Read the data from the Excel file into a pandas DataFrame
df = pd.read_excel('Z:\Compliance Reporting Team\Over Received ERS\Automation\Over Received Automation - Test 12-16.xlsm', sheet_name= 'Paste sheet', converters={'Purchasing Document':str,'Purchasing Doc Item':str})

# Get a list of buyer names
buyer_names = df['Name'].unique()

# Loop through each buyer name
for name in buyer_names:
  if pd.isnull(name) == True:
    continue
  # Filter the report data dataframe to only include rows with the current buyer name
  filtered_df = df[df['Name'] == name]

  #add a concat column so we can later pull comments
  filtered_df = filtered_df.assign(Concat = filtered_df['Purchasing Document'] + filtered_df["Purchasing Doc Item"])
  filtered_df['Concat'] = filtered_df['Concat'].astype(str)

  #Define the files we are writing to (use try to create file if none present)
  buyer_path = 'Z:\Compliance Reporting Team\Over Received ERS\Automation\data\\' + name + ".xlsx"
  new_condition = 1 
  try:
   pd.read_excel(buyer_path)
   new_condition = 0
  except FileNotFoundError:
     filtered_df.to_excel(buyer_path, index=False)
     
  #add a comments column to new data
  filtered_df = filtered_df.assign(Comments =None) 

  #Save the current(old) data in the buyer sheet and add concat column
  if new_condition == 0:
    old_data = pd.read_excel(buyer_path, usecols="C,D,G", converters={'Purchasing Document':str,'Purchasing Doc Item':str})
    old_data = old_data.assign(Concat = old_data['Purchasing Document'] + old_data["Purchasing Doc Item"])
    old_data['Concat'] = old_data['Concat'].astype(str)
    old_data.drop(old_data.columns[[0,1]],axis = 1, inplace= True)
  #use a join to pull in comments from old_data into filtered_df on concat
    filtered_df = filtered_df.drop('Comments', axis = 1)
    filtered_df = filtered_df.merge(old_data, on= ['Concat'], how='left')
  # Write the filtered DataFrame to a Excel file with the name of the buyer
  
  filtered_df = filtered_df.drop('Concat', axis = 1)
  filtered_df.to_excel(buyer_path, index=False)
 
   
 
