import glob
import pandas as pd
import bizdays
import os, sys
import functions
import constants as c
from numpy import array, array_equal

def create_dataframe(month, year, path):
    # Build list of files for the chosen month.
    filepath = f'{path}{month} {year}/*.XLS'
    filelist = glob.glob(filepath)

    # If an error log from a previous running of the report exists, delete it.
    errorlog = f'{path}/Error Logs/{month} {year} error log.txt'
    try:
        os.remove(errorlog)
    except:
        pass

    # Check if folder is empty. If it is, there is no reason to proceed.
    if len(filelist) == 0:
        print(f'There are currently no cards processed for {month} {year}.')
        sys.exit()

    # Initialize a calendar that will calculate the number of business days between two dates.
    cal = bizdays.Calendar(c.holidays, ['Sunday', 'Saturday'])

    # Create empty dataframe that will contain all card data.
    cards = pd.DataFrame()

    # Read each Excel file into the cards dataframe.
    for filename in filelist:
        data = pd.read_excel(filename)
        error = False 
        # Check to see if spreadsheet has the correct structure for reporting.
        # If it does not, add an entry to the error log for further investigation.
        if (len(data.columns.values) != len(c.sheet_columns)) or not array_equal(data.columns.values[0:7], array(c.sheet_columns[0:7])):
            error = True
            log = functions.open_error_log(errorlog)
            log.write(f'{filename} is incompatible with tribute spreadsheets.\n')
        else:
            # Removes potential inconsistencies in manually entered column names.
            data.columns.values[7:] = c.sheet_columns[7:]
            # Isolate dates columns as they are the only ones important for calculations.
            dates = data[['Gift Date Added', 'Date Pulled', 'Date Sent']]
            # See if any of the dates are missing.
            nans = pd.isnull(dates).any(1)
            nans_rows = nans.index[nans == True]
            if len(nans_rows) != 0:
                error = True
                log = functions.open_error_log(errorlog)
                for row in nans_rows:
                    log.write(f'Row {row+2} in {filename} is missing one or more dates.\n')
            # Check if any out-of-range dates are present.
            pulled_vs_sent = data['Date Pulled'] > data['Date Sent']
            if pulled_vs_sent.any():
                error = True
                log = functions.open_error_log(errorlog)
                indexes = pulled_vs_sent.index[pulled_vs_sent].tolist()
                for index in indexes:
                    log.write(f'Row {index+2} in {filename} has a Date Sent earlier than the Date Pulled\n')

        if error == False:
            # If no errors were found, add contents of spreadsheet to the dataframe.       
            cards = pd.concat([cards, data], ignore_index=True)
    # If there were any errors found, end process so user can investigate and correct issues.
    try:
        if log.mode == 'a+':
            print('There was a problem generating the report.  Please check error log.')
            log.close
            sys.exit()
    except NameError:
        # Error log was never opened, so there were no errors found.  Proceed with report generation.
        pass

    # Convert date columns to strings. Necessary to calculate differences in business days.  
    cards['Gift Date Added'] = cards['Gift Date Added'].astype(str)
    cards['Date Pulled'] = cards['Date Pulled'].astype(str)
    cards['Date Sent'] = cards['Date Sent'].astype(str)

    # Convert Constituent ID column to string to improve readability of final spreadsheet.
    cards['Constituent ID'] = cards['Constituent ID'].astype(str)

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
    return (cards, pulled, sent)