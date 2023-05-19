#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np


# In[ ]:


# Load the Excel file
file_path = '/Users/LarryZhang/Desktop/Randomization.xlsx'
xl = pd.ExcelFile(file_path)


# In[ ]:


# Sheets to read and number of samples from each
sheets_and_samples = {
    'treat_ipad': 200,
    'control_ipad': 200,
    'treat_ac': 35,
    'control_ac': 35,
}


# In[ ]:


# Create an empty DataFrame to store samples
samples = pd.DataFrame()


# In[ ]:


# Loop over each sheet
for sheet, n_samples in sheets_and_samples.items():
    # Read the sheet into a DataFrame
    df = xl.parse(sheet)
    
    # Check if the sheet has enough rows
    if len(df) < n_samples:
        print(f"Sheet {sheet} has less than {n_samples} rows.")
        continue

    # Randomly sample rows
    sample_df = df.sample(n_samples, random_state=77)

    # Add a new column to indicate the source of the sample
    sample_df['source'] = sheet

    # Append to the samples DataFrame
    samples = samples.append(sample_df)


# In[ ]:


# Write the samples DataFrame to a new Excel file
output_file_path = '/Users/LarryZhang/Desktop/Randomization_result.xlsx'
samples.to_excel(output_file_path, index=False)

