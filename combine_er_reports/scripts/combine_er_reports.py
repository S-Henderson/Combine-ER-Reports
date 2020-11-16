"""
Author: Scott Henderson
Last Updated: Nov 16, 2020

Purpose: Combine weekend ER reports to save time and potential manual copy/paste mistakes

Input: Raw ER reports (.xls) from project 'data/raw' folder
Output: Combined (appended) ER reports (.xlsx) into 'data/exports' folder
"""

import os
import pandas as pd
import numpy as np
import openpyxl
import glob
from datetime import datetime, timedelta
import pandas.io.formats.excel

# Remove pandas excel header formatting
pandas.io.formats.excel.ExcelFormatter.header_style = None

#--------------- PURPOSE ---------------#

print("Purpose: Combine weekend ER reports to save time and potential manual copy/paste mistakes")

print("----------------------------------------------------------------------------------------------------")

#--------------- ASCII ART ---------------#

print(r"""
_________               ___.   .__                _____________________  __________                             __          
\_   ___ \  ____   _____\_ |__ |__| ____   ____   \_   _____/\______   \ \______   \ ____ ______   ____________/  |_  ______
/    \  \/ /  _ \ /     \| __ \|  |/    \_/ __ \   |    __)_  |       _/  |       _// __ \\____ \ /  _ \_  __ \   __\/  ___/
\     \___(  <_> )  Y Y  \ \_\ \  |   |  \  ___/   |        \ |    |   \  |    |   \  ___/|  |_> >  <_> )  | \/|  |  \___ \ 
 \______  /\____/|__|_|  /___  /__|___|  /\___  > /_______  / |____|_  /  |____|_  /\___  >   __/ \____/|__|   |__| /____  >
        \/             \/    \/        \/     \/          \/         \/          \/     \/|__|                           \/ 
""")

print("----------------------------------------------------------------------------------------------------")

#--------------- SOURCE AND DESTINATION PATHS ---------------#

# Source directory where ER reports are saved
src_dir = os.path.join(os.path.expanduser("~"), "Desktop", "python_projects", "combine_er_reports", "data", "raw")

# Destination directory where ER reports are exported to
dst_dir = os.path.join(os.path.expanduser("~"), "Desktop", "python_projects", "combine_er_reports", "data", "exports")

#--------------- LIST FILES ---------------#

# Get basename of each file
file_list = [os.path.basename(file) for file in glob.glob(f"{src_dir}/Fraud Results for*.xls")]

print(*file_list, sep='\n')

num_of_files = len(file_list)
print(f"There are {num_of_files} files")

print("----------------------------------------------------------------------------------------------------")

#--------------- FIND UNIQUE CLIENTS ---------------#

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

print("----------------------------------------------------------------------------------------------------")

#--------------- DATAFRAME PREP ---------------#

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

#--------------- CHANGE WORKING DIRECTORY ---------------#

# To save combined ER files to exports folder
os.chdir(dst_dir)

#--------------- APPENDING & SAVING FILES ---------------#

def append_files():
    """
    Takes client -> creates empty df for client -> finds all files matching that client and appends data to empty df -> moves to next client match
    """
    
    # Create blank df for each client
    for client in client_list:
        
        print(client)
        
        # Find the files associated with each client
        files = glob.glob(f"{src_dir}/*{str(client)}*.xls")
        
        # Optional -> print list of client group of files
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
            
            # Optional -> print each df to check
            #print(df_er_file.head(5))
            
            # Append data to client blank df
            df_blank = df_blank.append(df_er_file, 
                                       ignore_index = True)
            
            # Set filename
            
            # For Mondays - add together weekend + Monday (past 3 days) & modified date is that Friday-Sat-Sun (past 3 days minus 1 etc)
            
            # Month/Year
            month = datetime.datetime.now().strftime('%m')
            year = datetime.datetime.now().strftime('%Y')
            
            # Days
            d = datetime.datetime.now().strftime('%d')
            d1 = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%d')
            d2 = (datetime.datetime.now() - datetime.timedelta(days=2)).strftime('%d')
            d3 = (datetime.datetime.now() - datetime.timedelta(days=3)).strftime('%d')

            # Filename
            filename = f"Fraud Results for {client} P-Date {month}.{d2}.{d1}.{d}-{year}_M-Date {month}.{d3}.{d2}.{d1}-{year}.xlsx"
            
            # Edit-able filename
            #filename = f"Fraud Results for {client} P-Date 11.07.08.09-2020_M-Date 11.06.07.08-2020.xlsx"
            
            print(f"Appending to file -> {filename}")
            
            # Write data
            writer = pd.ExcelWriter(filename, 
                                    engine = "openpyxl")
            
            df_blank.to_excel(writer, 
                              sheet_name = "AllSessionSorted",
                              index = False)
            
            print("Exporting file...")
            
            writer.save()
            
            print(f"Successfully exported file for: {client}")
            
            print("----------------------------------------------------------------------------------------------------")

# Call loop function
append_files()

#--------------- SCRIPT COMPLETED ---------------#

print("----------------------------------------------------------------------------------------------------")

print("Script Successfully Completed")

input("Press Enter to Continue...")