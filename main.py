# ---------------------------------------------------------
# Import Libraries
# ---------------------------------------------------------

import numpy as np  # for array- and numerical data handling.
import pandas as pd  # for managing data within dataframes.
pd.set_option('mode.chained_assignment', None)  # Allow for chain assignments.
# For database connectivity
import pyodbc
from sqlalchemy import create_engine
import time
# For accessing Outlook via Python
import win32com.client as win32

"""
In the case where data is available from a SQL database:

- Establish a connection to the database using pyodbc.
- Fetch data with a SQL query.

"""

# ---------------------------------------------------------
# Setup Connection
# ---------------------------------------------------------

def SETUP_CONNECTION(DRIVER, SERVER, DATABASE, USERNAME, PASSWORD):
    
    # Setup your database connection here using the relevant server and database details:
    CONNECTION = pyodbc.connect(
        'DRIVER=' + DRIVER + ';SERVER=' + SERVER + ';DATABASE=' + DATABASE + ';UID=' + USERNAME + ';PWD=' + PASSWORD + ';Authentication=ActiveDirectoryPASSWORD')

    return CONNECTION
  
# ---------------------------------------------------------
# Fetch data from database
# ---------------------------------------------------------

def GET_DATA(CONNECTION):

    print('\n\n')
    print('Fetching data...')
    print('\n\n')
    
    # SQL Query for fetching data from database.
    QUERY = """
    SELECT * FROM <TABLE>
    """

    DF = pd.read_sql_query(QUERY, con=CONNECTION)
    print('Success')

    return DF

'''
If you have data from a local database, run these:
'''
#CONNECTION = SETUP_CONNECTION()
#data = GET_DATA(CONNECTION)

# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# Demonstration
# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

# Generate sample data to show functionality of this program.
data = pd.DataFrame({'Id': [1,2,3,4,5], 'Check': [True, False, True, False, False] })

# ---------------------------------------------------------
# Generate email body using HTML. Specify outbox.
# ---------------------------------------------------------

class GENERATE_EMAIL:
    
    def __init__(self, ID):
        
        self.ID = ID
        
        self.email_address = GENERATE_EMAIL.get_email_address(self)[0]
        self.cc = GENERATE_EMAIL.get_email_address(self)[1]
        
        self.subject = GENERATE_EMAIL.get_subject(self)
        
        self.email_body = GENERATE_EMAIL.get_body(self)
        
    
    def get_id(self):
        return self.ID
    
    
    def get_email_address(self):

        self.email_address = "YOUR_EMAIL_ADDRESS@DOMAIN.COM"
        self.cc = "ADDRESS1@DOMAIN.COM; ADDRESS2@DOMAIN.COM"

        return self.email_address, self.cc
    
    def get_subject(self):
        
        self.subject = "EMAIL SUBJECT"
        
        return self.subject
    
    def get_body(self):
            
        self.email_body = html = """\n
        <html>
          <head></head>
          <body>
            INSERT YOUR HTML CODE
            </p>
          </body>
        </html>
        """
        return self.email_body

# ---------------------------------------------------------
# Send email with win32
# ---------------------------------------------------------
            
def SEND(info):
    
    #outlook = win32.Dispatch('outlook.application')
    #mail = outlook.CreateItem(0)
    #mail.To = info.email_address
    #mail.CC = info.cc
    #mail.Subject = info.subject
    #mail.HTMLBody = info.email_body
    
    # Uncomment to enable functionality
    #mail.Send()
    
    print('EMAIL SENT!')
    
    return None
  
# ---------------------------------------------------------
# Send email based on some condition within the dataset.
# ---------------------------------------------------------

start_time = time.time()
for i in range(len(data)):
    # Some condition
    if data['Check'][i] == True:
        EMAIL = GENERATE_EMAIL(data['Id'][i])
        # Send email
        SEND(EMAIL)
        print("--- %s seconds ---" % (time.time() - start_time))
