import openpyxl

# import AreaChart class from openpyxl.chart sub_module
from openpyxl.chart import AreaChart,Reference


wb = openpyxl.Workbook()
sheet = wb.active

# write o to 9 in 1st column of the active sheet
for i in range(10):
	sheet.append([i])

values = Reference(sheet, min_col = 1, min_row = 1,
						max_col = 1, max_row = 10)

# create object of AreaChart class
chart = AreaChart()

chart.add_data(values)

# set the title of the chart
chart.title = " AREA-CHART "

# set the title of the x-axis
chart.x_axis.title = " X-AXIS "

# set the title of the y-axis
chart.y_axis.title = " Y-AXIS "

# add chart to the sheet
# the top-left corner of a chart
# is anchored to cell E2 .
sheet.add_chart(chart, "E2")

# save the file
wb.save("AreaChart.xlsx")
