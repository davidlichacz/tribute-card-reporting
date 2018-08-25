from loaddata import create_dataframe
from monthlyreport import summary_stats, open_file

if __name__ == '__main__':
	# Get month and year of report from user.  Ultimately this will be a selection from a drop-down list
    # so error handling is omitted for now.
	month = input('Enter month: ')
	year = input('Enter year: ')
	path = '/Users/davidlichacz/Tribute Spreadsheets/'
	cards, pulled, sent = create_dataframe(month, year, path)
	report = summary_stats(cards, pulled, sent, month, year, path)
	open_file(report)