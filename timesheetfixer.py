# Timesheet Import Fixer For Python 3.5
# Because the bookeeper will destroy the database if they attempt to use PostgreSQL
# Global imports
import csv
import time
import datetime
import dateutil.parser as dparser
from collections import defaultdict
# Global variables
year = int()
month = int()
day = int()

# Main Process
def main():
	print ("--- Welcome to Timesheet Import Fixer ---\n")
	print ("Please download the raw timesheet file from Reports > Employee > Timesheet\n--- Make sure it is sorted by Employee\n--- Make sure the date range is from pay period start to pay period end\n--- Rename the file to pay period start date: 'Timesheet-MM-DD-YYYY.csv'\n--- Also place into a folder with same start date: 'MM-DD-YYYY'\n")
	print ("\nA new file will be created with 'Edited' appended onto the front in each folder.\n--- If there are any errors the document will store them in a seperate column\n--- If changes need to be made do so to the original file and run this script again.\n--- Verify that it is still saved as a CSV file after saving if changes needed to be made\n")
	input ("Press enter when ready to continue...\n")
	timesheetparser()

def timesheetparser():
	# Imports
	import glob
	import os
	# Variables
	global month
	global day
	global year
	global timesheet
	global current_timesheet_location
	# Run for Salary or Hourly
	print ("Operations can be performed on the following dictionaries:\n")
	salarytimesheets = [os.path.basename(x) for x in glob.glob('./salary/*/Timesheet-*.csv')]
	hourlytimesheets = [os.path.basename(x) for x in glob.glob('./hourly/*/Timesheet-*.csv')]
	gttimesheets = [os.path.basename(x) for x in glob.glob('./gt/*/Timesheet-*.csv')]
	print (sorted(salarytimesheets))
	print (sorted(hourlytimesheets))
	print (sorted(gttimesheets))
	input ("\nPress enter to continue: ")
	# Time Counter
	total_start_time = time.time()
	for timesheet in salarytimesheets:
		current_timesheet_location = glob.glob('./salary/*/'+timesheet)
		date = dparser.parse(timesheet,fuzzy=True)
		year = date.year
		month = date.month
		day = date.day
		print ('\n')
		salary = True
		gt = False
		hourly = False
		mrclean(salary, gt, hourly)
		print ("\nEdits complete for Salary:\n")
		print (salarytimesheets)
	for timesheet in hourlytimesheets:
		current_timesheet_location = glob.glob('./hourly/*/'+timesheet)
		date = dparser.parse(timesheet,fuzzy=True)
		year = date.year
		month = date.month
		day = date.day
		print ('\n')
		salary = False
		gt = False
		hourly = True
		mrclean(salary, gt, hourly)
		print ("\nEdits complete for Hourly:\n")
		print (hourlytimesheets)
	for timesheet in gttimesheets:
		current_timesheet_location = glob.glob('./gt/*/'+timesheet)
		date = dparser.parse(timesheet,fuzzy=True)
		year = date.year
		month = date.month
		day = date.day
		print ('\n')
		salary = False
		gt = True
		hourly = False
		mrclean(salary, gt, hourly)
		print ("\nEdits complete for Weekly:\n")
		print (gttimesheets)
	print ("\nTotal time\n--- {} seconds ---\n".format(time.time() - total_start_time))

def mrclean(salary, gt, hourly):
	# imports
	import itertools
	# Variables
	week1 = defaultdict(int)
	week2 = defaultdict(int)
	start_time = time.time()
	# For Headers
	paydate_start = datetime.date(year, month, day)
	paytime_start = datetime.datetime(year, month, day)
	if salary:
		print (day)
		print (month)
		paydate_half = paydate_start + datetime.timedelta(days=8)
		paydate_end = paydate_half + datetime.timedelta(days=7)
		paytime_half = paytime_start + datetime.timedelta(days=8)
		paytime_end = paytime_half + datetime.timedelta(days=8)
		g = open('./salary/'+str(month)+'-'+str(day)+'-'+str(year)+'/Edited-Salary'+timesheet, 'wt')
	if gt:
		paydate_end = paydate_start + datetime.timedelta(days=8)
		paytime_end = paytime_start + datetime.timedelta(days=8)
		g = open('./gt/'+str(month)+'-'+str(day)+'-'+str(year)+'/Edited-GT'+timesheet, 'wt')
	if hourly:
		paydate_half = paydate_start + datetime.timedelta(days=7)
		paydate_end = paydate_half + datetime.timedelta(days=6)
		paytime_half = paytime_start + datetime.timedelta(days=7)
		paytime_end = paytime_half + datetime.timedelta(days=7)
		g = open('./hourly/'+str(month)+'-'+str(day)+'-'+str(year)+'/Edited-'+timesheet, 'wt')
	# CSV Opening/Reading/Writing
	f = open(current_timesheet_location[0], 'rt')
	f2 = open('assets/employee_type.csv', 'rt')
	csv_r = csv.reader(f)
	csv_r2 = csv.reader(f2)
	csv_w = csv.writer(g)
	# Salary vs Hourly dictionary Creation from CSV
	employeetype = dict((rows[0],rows[1]) for rows in csv_r2)
	# Skip header and add new header to new CSV
	if salary:
		print ("Starting Edits for Salary:\n"+timesheet+"\n---\nPay Period Start: " + str(paydate_start) + "\nPay Period Half: " + str(paydate_half) + "\nPay Period End: " + str(paydate_end) + "\n" + "---")
	if gt:
		print ("Starting Edits for Green Thumb:\n"+timesheet+"\n---\nPay Period Start: " + str(paydate_start) + "\nPay Period End: " + str(paydate_end) + "\n" + "---")
	if hourly:
		print ("Starting Edits for Hourly:\n"+timesheet+"\n---\nPay Period Start: " + str(paydate_start) + "\nPay Period Half: " + str(paydate_half) + "\nPay Period End: " + str(paydate_end) + "\n" + "---")
	csv_w.writerow(('Timesheet For:',str(paydate_start) + ' to ' + str(paydate_end)))
	csv_w.writerow(('Employee','Clock In', 'Clock Out', 'Hours'))
	next(f)
	# Setup double iteration for proper hours addition at end of each series of persons
	reader1,reader2 = itertools.tee(csv_r)
	next(reader2)
	for row,row_next in zip(reader1,reader2):
		# Variables to make code easier to read
		x = row[0]
		dateformat = ("%m/%d/%Y %I:%M %p")
		if any(field.strip() for field in row[1] and row[2]):
				# Time re-calculation and rounding to nearest 15 minutes
				tm1, tm2 = dparser.parse(row[1]), dparser.parse(row[2])
				tm1 += datetime.timedelta(minutes=8)
				tm1 -= datetime.timedelta(minutes=tm1.minute % 15, seconds=tm1.second, microseconds=tm1.microsecond)
				tm2 += datetime.timedelta(minutes=8)
				tm2 -= datetime.timedelta(minutes=tm2.minute % 15, seconds=tm2.second, microseconds=tm2.microsecond)
				delta = (tm2 - tm1)
				hours = (delta.total_seconds()/(3600))
				# Calculate PTO and Hours for Week 1
				if salary:
					if employeetype.get(x, 'default') == 'salary':
						if (paytime_start <= tm1 < paytime_half):
							if hours < 0:
								week1[row[0]+"-error"] += 1
								csv_w.writerow((x, row[1], row[2], delta, 'Time-Range-Error'))
								continue
							if week1[x]>40:
								week1[x+"-hours"] += hours
								week1[x+"-pto"] += (hours*2)
								if hours >= 10:
									csv_w.writerow((x, row[1], row[2], delta, 'Large-Time-Range'))
								else:
									csv_w.writerow((x, row[1], row[2], delta))
							if week1[x]<=40:
								week1[x+"-hours"] += hours
								week1[x+"-pto"] += hours
								if hours >= 10:
									csv_w.writerow((x, row[1], row[2], delta, 'Large-Time-Range'))
								else:
									csv_w.writerow((x, row[1], row[2], delta))
						# Calculate PTO and Hours for Week 2
						if (paytime_half <= tm1 < paytime_end):
							if hours < 0:
								week2[x+"-error"] += 1
								csv_w.writerow((x, row[1], row[2], delta, 'Time-Range-Error'))
								continue
							if week2[x]>40:
								week2[x+"-hours"] += hours
								week2[x+"-pto"] += (hours*2)
								if hours >= 10:
									csv_w.writerow((x, row[1], row[2], delta, 'Large-Time-Range'))
								else:
									csv_w.writerow((x, row[1], row[2], delta))
							if week2[x]<=40:
								week2[x+"-hours"] += hours
								week2[x+"-pto"] += hours
								if hours >= 10:
									csv_w.writerow((x, row[1], row[2], delta, 'Large-Time-Range'))
								else:
									csv_w.writerow((x, row[1], row[2], delta))
					else:
						continue
				if gt:
					if employeetype.get(x, 'default') == 'gt':
						if (paytime_start <= tm1 < paytime_end):
							if hours < 0:
								week1[row[0]+"-error"] += 1
								csv_w.writerow((x, row[1], row[2], delta, 'Time-Range-Error'))
								continue
							if week1[x]>40:
								week1[x+"-hours"] += hours
								week1[x+"-pto"] += (hours*2)
								if hours >= 10:
									csv_w.writerow((x, row[1], row[2], delta, 'Large-Time-Range'))
								else:
									csv_w.writerow((x, row[1], row[2], delta))
							if week1[x]<=40:
								week1[x+"-hours"] += hours
								week1[x+"-pto"] += hours
								if hours >= 10:
									csv_w.writerow((x, row[1], row[2], delta, 'Large-Time-Range'))
								else:
									csv_w.writerow((x, row[1], row[2], delta))
					else:
						continue
				if hourly:
					if employeetype.get(x, 'default') == 'default':
						if (paytime_start <= tm1 < paytime_half):
							if hours < 0:
								week1[row[0]+"-error"] += 1
								csv_w.writerow((x, row[1], row[2], delta, 'Time-Range-Error'))
								continue
							if week1[x]>40:
								week1[x+"-hours"] += hours
								week1[x+"-pto"] += (hours*2)
								if hours >= 10:
									csv_w.writerow((x, row[1], row[2], delta, 'Large-Time-Range'))
								else:
									csv_w.writerow((x, row[1], row[2], delta))
							if week1[x]<=40:
								week1[x+"-hours"] += hours
								week1[x+"-pto"] += hours
								if hours >= 10:
									csv_w.writerow((x, row[1], row[2], delta, 'Large-Time-Range'))
								else:
									csv_w.writerow((x, row[1], row[2], delta))
						# Calculate PTO and Hours for Week 2
						if (paytime_half <= tm1 < paytime_end):
							if hours < 0:
								week2[x+"-error"] += 1
								csv_w.writerow((x, row[1], row[2], delta, 'Time-Range-Error'))
								continue
							if week2[x]>40:
								week2[x+"-hours"] += hours
								week2[x+"-pto"] += (hours*2)
								if hours >= 10:
									csv_w.writerow((x, row[1], row[2], delta, 'Large-Time-Range'))
								else:
									csv_w.writerow((x, row[1], row[2], delta))
							if week2[x]<=40:
								week2[x+"-hours"] += hours
								week2[x+"-pto"] += hours
								if hours >= 10:
									csv_w.writerow((x, row[1], row[2], delta, 'Large-Time-Range'))
								else:
									csv_w.writerow((x, row[1], row[2], delta))
					else:
						continue
				# Used if Error with Dates
				if (tm1 < paytime_start or tm1 >= paytime_end):
					week1[x+"-error"] += 1
					csv_w.writerow((x, row[1], row[2], delta, 'Date-Range-Error'))
					print ("Row # %s has a wrong date or the time was entered wrong '%s'" % (csv_r.line_num, (x,row[1],row[2])))
				# Used to append new data at end of name series
				if row_next[0] == x or x == str():
					continue
				elif row_next[0] != x:
					# Add up PTO before writing
					total_pto = {x: week1.get(x+'-pto', 0) + week2.get(x+'-pto', 0) for x in set(week1).union(week2)}
					total_hours = {x: week1.get(x+'-hours', 0) + week2.get(x+'-hours', 0) for x in set(week1).union(week2)}
					# Calculate Overtime
					if week1[x+"-hours"]-40 > 0:
						week1[x+"-overtime"] += week1[x+"-hours"]-40
					if week2[x+"-hours"]-40 > 0:
						week2[x+"-overtime"] += week2[x+"-hours"]-40
					total_overtime = {x: week1.get(x+'-overtime', 0) + week2.get(x+'-overtime', 0) for x in set(week1).union(week2)}
					errors = {x: week1.get(x+'-error', 0) + week2.get(x+'-error', 0) for x in set(week1).union(week2)}
					# Conversion of Hours to PTO
					total_pto.update((x, round(y*0.04, 2)) for x, y in total_pto.items())
					# Add up Hours and total hours before writing
					if errors[x] != 0:
						csv_w.writerows([[' '],['Total Hours','Week 1 Hours','Week 2 Hours','PTO','Errors'],[total_hours[x],week1[x+"-hours"],week2[x+"-hours"],total_pto[x],errors[x]],[' ']])
					else:
						if employeetype.get(x, 'default') != 'default':
							csv_w.writerows([[' '],['Total Hours','Week 1 Hours','Week 2 Hours','PTO'],[total_hours[x],week1[x+"-hours"],week2[x+"-hours"],total_pto[x]],[' '],['Signature'],[' ']])
						else:
							if total_overtime[x] > 0:
								csv_w.writerows([[' '],['Total Hours','Week 1 Hours','Week 2 Hours'],[total_hours[x],week1[x+"-hours"],week2[x+"-hours"]],['Total Overtime','Week 1 Overtime','Week 2 Overtime'],[total_overtime[x],week1[x+"-overtime"],week2[x+"-overtime"]],[' '],['Signature'],[' ']])
							else:
								csv_w.writerows([[' '],['Total Hours','Week 1 Hours','Week 2 Hours'],[total_hours[x],week1[x+"-hours"],week2[x+"-hours"]],[' '],['Signature'],[' ']])
		elif any(field.strip() for field in row[1] or row[2]):
				# Column for missing info
				week1[x+"-error"] += 1
				csv_w.writerow((x, row[1], row[2], ' ', 'This row needs data'))
				print ("Row # %s has missing information '%s'\n") % (csv_r.line_num, (x,row[1],row[2]))
	print ("\nEdit complete and named Edited-"+timesheet+"\n--- {} seconds ---\n".format(time.time() - start_time))
	f.close()
	f2.close()
	g.close()

# Start Function
main()
# End Function
exit()
