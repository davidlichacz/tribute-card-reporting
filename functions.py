def bizdays_neg(calendar, date1, date2):
	try:
		return calendar.bizdays(date1, date2)
	except ValueError:
		return -calendar.bizdays(date2, date1)