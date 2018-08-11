def bizdays_neg(calendar, date1, date2):
	try:
		return calendar.bizdays(date1, date2)
	except ValueError:
		return -calendar.bizdays(date2, date1)

def as_text(value):
    if value is None:
        return ""
    return str(value)

def adjust_width(sheet):
	for column_cells in sheet.columns:
		length = max(len(as_text(cell.value)) for cell in column_cells)
		sheet.column_dimensions[column_cells[0].column].width = length