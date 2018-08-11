import glob
import pandas as pd
import bizdays
import os
import sys
import openpyxl
import subprocess, os
import functions
from statholidays import holidays


# Get month and year of report from user.  Ultimately this will be a selection from a drop-down list
# so error handling is omitted for now.
month = input('Enter month: ')
year = input('Enter year: ')

# Build list of files for the chosen month.
path = '/Users/davidlichacz/Tribute Spreadsheets/'
filepath = f'{path}{month} {year}/*.XLS'
filelist = glob.glob(filepath)

# Check if folder is empty.
if len(filelist) == 0:
    print(f'There are currently no cards processed for {month} {year}.')
    sys.exit()

# Initialize a calendar that will calculate the number of business days between two dates.
cal = bizdays.Calendar(holidays, ['Sunday', 'Saturday'])

# Create empty dataframe that will contain all card data.
cards = pd.DataFrame()

# Read each Excel file into the cards dataframe.
for filename in filelist:
    data = pd.read_excel(filename)
    cards = pd.concat([cards, data], ignore_index=True)

# Convert date columns to strings. Necessary to calculate differences in business days.  
cards['Gift Date Added'] = cards['Gift Date Added'].astype(str)
cards['Date Pulled'] = cards['Date Pulled'].astype(str)
cards['Date Sent'] = cards['Date Sent'].astype(str)

# Calculate differences in business days and insert them into dataframe.
rows = cards.shape[0]
pulled = []
sent = []

for k in (range(rows)):
    pulled.append(functions.bizdays_neg(cal, cards['Gift Date Added'][k], cards['Date Pulled'][k]))
    sent.append(functions.bizdays_neg(cal, cards['Gift Date Added'][k], cards['Date Sent'][k]))

cards.insert(loc=9, column='Business Days Until Pulled', value=pulled)
cards.insert(loc=10, column='Business Days Until Sent', value=sent)

# Sort dataframe for logical readability.
cards = cards.sort_values(by=['Gift Date Added', 'Tribute Card Type'])

# Prepare, save and open final file.
reportfile = f'{path}{month} {year}/{month} {year} Tribute Cards.xlsx'

writer = pd.ExcelWriter(reportfile)

cards.to_excel(writer, index=False)
writer.save()

if sys.platform.startswith('darwin'):
    subprocess.call(('open', reportfile))
elif os.name == 'nt':
    os.startfile(reportfile)
elif os.name == 'posix':
    subprocess.call(('xdg-open', reportfile))