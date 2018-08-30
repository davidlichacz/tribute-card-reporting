from loaddata import create_dataframe
from monthlyreport import summary_stats, open_file
from annualreport import annual_report
from constants import months


if __name__ == '__main__':
	# Get month and year of report from user.  Ultimately this will be a selection from a drop-down list
    # so error handling is omitted for now.
	month = input('Enter month: ')
	year = input('Enter year: ')
	# Define range for fiscal year.
	if month in months[0:9]:
		fiscal = f'{year}-{int(year)+1}'
	else:
		fiscal = f'{int(year)-1}-{year}'
	path = '/Users/davidlichacz/Tribute Spreadsheets/'
	cards, pulled, sent = create_dataframe(month, year, path)
	report = summary_stats(cards, pulled, sent, month, year, path)
	annual_report(month, fiscal, path, report[1])
	#open_file(report[0])