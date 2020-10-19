"""
Author: Scott Henderson
Last Updated: Oct 18, 2020

Purpose: Combine weekend ER reports to save time and potential manual copy/paste mistakes

Input: Raw ER reports (.xls) into project 'data/raw' folder
Output: Combined ER reports (.xlsx) into 'data/exports' folder
"""

import os

import pandas as pd
import numpy as np
import glob
import datetime

import pandas.io.formats.excel

# Remove pandas excel header formatting
pandas.io.formats.excel.ExcelFormatter.header_style = None

# Source directory where ER reports are saved
src_dir = os.path.join(os.environ['USERPROFILE'], "Desktop", "python_projects", "combine_er_reports", "data", "raw")

# Get basename of each file
file_list = [os.path.basename(x) for x in glob.glob(f"{src_dir}/Fraud Results for*.xls")]

print(*file_list, sep='\n')

num_of_files = len(file_list)
print(f"There are {num_of_files} files")

# Empty list to store unique client names
client_list =[]

# Split file_name, grab client name, append unique values to list
for filename in file_list:
    
    # String split
    type = filename.split(" ")
    
    # Grab 4th word -> client name
    client = type[3]
    
    # Check if exists in client_list
    if client not in client_list: 
        client_list.append(client) 
		
		
# List of unique client names
print(*client_list, sep='\n')

num_of_clients = len(client_list)
print(f"There are {num_of_clients} unique clients")

# Column headers for ER file
columns = ["Session", 
           "Client Code", 
           "Module", 
           "First Name", 
           "Last Name", 
           "Address", 
           "City", 
           "State", 
           "Zip", 
           "Email", 
           "Dealer", 
           "SPCode", 
           "Invoice", 
           "Brand", 
           "Product Line", 
           "Models", 
           "Model QTY", 
           "SerialNumber", 
           "SPDescription", 
           "Process ID", 
           "LevSessionNumber", 
           "Status", 
           "Date of Sale", 
           "Created Date", 
           "Modified Date", 
           "Lev Score", 
           "Comments", 
           "Patient First Name", 
           "Patient Last Name"]

# Destination directory where ER reports are exported to
dst_dir = os.path.join(os.environ['USERPROFILE'], "Desktop", "python_projects", "combine_er_reports", "data", "exports")
os.chdir(dst_dir)

# Create blank df for each client
for client in client_list:
    
    print(client)
    
    # Find the files associated with each client
    files = glob.glob(f"{src_dir}/*{str(client)}*.xls")
    
    #print(files)
    
    # Create a blank dataframe to store each client's data
    print("Creating blank dataframe...")
    
    # Create df with column headers
    df_blank = pd.DataFrame(columns = columns)
    
    # Appends all the files for the relevant client into one df and saves it 
    for file in files:
        
        print("Reading file...")
        
        df_er_file = pd.read_excel(file,
                                   sheet_name = "AllSessionSorted")
        
        print("Sucessfully read file...")
        
        #print(df_er_file.head(5))
        
        # Append data to client blank df
        df_blank = df_blank.append(df_er_file, 
                                   ignore_index = True)
        
        # Set filename -> need to update P-Date & M-Date each time
        #filename = "Fraud Results for " + client + " P-Date 10-10.11.12.13-2020_M-Date 10-09.10.11.12-2020" + ".xlsx"
        
        filename = f"Fraud Results for {client} P-Date 10-10.11.12.13-2020_M-Date 10-09.10.11.12-2020.xlsx"
        
        #print(filename)
        
        # Write data
        writer = pd.ExcelWriter(filename, 
                                engine = "openpyxl")
        
        df_blank.to_excel(writer, 
                          sheet_name = "AllSessionSorted",
                          index = False)
        
        print("Exporting file...")
        
        writer.save()
        
        print(f"Successfully exported file for: {client}")
        
        print("****************************************************************************************************")
        


#writer.close()

print("Script Successfully Completed")

input("Press Enter to Continue...")