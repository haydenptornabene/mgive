# This script will take .xlsx (Excel) files and peform the following operation:
# 1. It will take a phone number written in some strange way, (303)-990-1002, and return
#    a string of that phone number, 3039901002.
# 2. This script is smart enough that it will take any column in the excel file titled 
#    'Phone Number', 'Mobile Number', 'Telephone', 'Contact Number', or something simialr.
#     If you an error results, make sure your column is labeld Phone Number
# 3. The script will return a fully processed .xlsx file

# WHEN RUNNING THIS SCRIPT, PUT THE FOLLOWING COMMAND INTO THE COMMAND LINE


#                python phone.massage.py InputFile.csv OutPutFile.xlsx
#        where InputFile.xlsx is your input file, and OutPutFile.xlsx your ouput

import xlrd
import re
from xlrd import open_workbook, cellname
import csv
import string
import sys
import itertools
import openpyxl
import glob
from pyexcel.cookbook import merge_all_to_a_book

###################################### PREAMBLE ABOVE #############################

# This will take the second to last argument in the command line, the input file name
filein = sys.argv[-2]
# This will take the last argument in the command line, the output file name
fileout = sys.argv[-1] 

# Open the desired file
book = open_workbook(filein)
sheet = book.sheet_by_index(0)

# Read header values into the list
keys = [sheet.cell(0, col_index).value for col_index in xrange(sheet.ncols)]

# We will use this list to find the column 'Phone Number' in the dictionary 
# which will serve as the identifier for later. 

# Create a dictionary with every value from the spreadsheet
dict_list_raw = []
for row_index in xrange(1, sheet.nrows):
    d = {keys[col_index]: sheet.cell(row_index, col_index).value 
         for col_index in xrange(sheet.ncols)}
    dict_list_raw.append(d)


# Remove duplicate dictionaries in dict_list_raw to create dict_list

dict_list = []
for x in dict_list_raw:
    if x not in dict_list:
        dict_list.append(x)

# Not every spreadsheet has "Mobile Number" as the header. The following is a list of possible entries:
header_sample = ['Mobile Number', 'Mobile Numbers', 'Phone Number', 'Phone Numbers', 'Telephone','Telephone Number', 'Telephone Numbers', 'Tele', 'Contact', 'Contact Number', 'Mobile Phone Number', 'Mobile Number']

# Set up interesection tool that compares the lists keys with header_sample
# The variable 'common' is equivalent to the header for telephone numbers
common = (list(set(keys) & set(header_sample)))
# Make header a string and then into an integer 
common = map(str, common) 

# A Dictionary entry is read as a [key:value]. We want all values with the key that denotes phone numbers

# Initialize the list
phone_numbers = []

# If there is a header match above, map that string as the key to each dictionary pair
if len(common) == 1:
	# Turn list into an single string value
	common = common[0]
	#Map header to values in dictionary
	phone_numbers = map (lambda x:x[common],dict_list)
else:
	print 'Please Change the header of you phone numbers to Mobile Number'

# The following chunk cleans up the phone numbers

# Convert from unicode to strings
phone_numbers = map(str, phone_numbers)

# THIS AREA COULD BE REWRITTEN WITH A FUNCTION BUT FOR NOW, IT FUNCTIONS

# For loop runs through phone number list and replaces - in phone numbers
for idx in range(len(phone_numbers)):
	phone_numbers[idx] = phone_numbers[idx].replace('-', '').replace(' ', '') 
# For loop runs through phone number list and replaces (,),*, and spaces in phone numbers
for idx in range(len(phone_numbers)):
	phone_numbers[idx] = phone_numbers[idx].replace('(', '').replace(' ', '') 
for idx in range(len(phone_numbers)):
	phone_numbers[idx] = phone_numbers[idx].replace(')', '').replace(' ', '') 
for idx in range(len(phone_numbers)):
	phone_numbers[idx] = phone_numbers[idx].replace('*', '').replace(' ', '') 
for idx in range(len(phone_numbers)):
	phone_numbers[idx] = phone_numbers[idx].replace('?', '').replace(' ', '') 
for idx in range(len(phone_numbers)):
	phone_numbers[idx] = phone_numbers[idx].replace('!', '').replace(' ', '')
for idx in range(len(phone_numbers)):
	phone_numbers[idx] = phone_numbers[idx].replace('@', '').replace(' ', '')  	
for idx in range(len(phone_numbers)):
	phone_numbers[idx] = phone_numbers[idx].replace('~', '').replace(' ', '')
for idx in range(len(phone_numbers)):
	phone_numbers[idx] = phone_numbers[idx].replace(' ', '').replace(' ', '') 

# Insert the header into the list
phone_numbers.insert(0, common)

# Remove spaces and the word phone in the header for this list
phone_numbers = [w.replace(' ', '') for w in phone_numbers]
phone_numbers = [w.replace('Phone', '') for w in phone_numbers]

# This chunk removes the awkward space in the MESSAGE value between 'is' and the date
for myitem in dict_list:
	for key in myitem:
		if key == 'MESSAGE':
			myitem[key] = myitem[key].replace("is  ", "is ")

# We need to remove the weird acute accent symbol in the 'you'll' message string. 
# It's not actually an apostrophe in the xslx file. First define the character's unicode 

symbolunicode = u'\u2019'

for myitem in dict_list:
	for key in myitem:
		if key == 'MESSAGE':
			myitem[key] = myitem[key].replace(symbolunicode, "")

# There is a weird dash everyone once in a while, this will remove it if it exists. 

symbolunicodedash = u'\u2013'

for myitem in dict_list:
	for key in myitem:
		if key == 'MESSAGE':
			myitem[key] = myitem[key].replace(symbolunicodedash, "")

# We need to insert the rest of the data into the csv while perserving order.
# Remove the phone number header from the keys list.

stringkeys = map(str, keys) 
stringkeys.remove(common)

# We need to populate a series of lists with the rest of the data in dict_list that will be written to the csv
# This next piece builds other columns (non phone number columns)

# Initialize the array datacolumns
datacolumns = []
# Stringkeys is a list containing the name of all columns other than the phone number header
for column in stringkeys:
  column_header_and_values = []
  column_header_and_values.append(column)
  for item in dict_list:
    column_header_and_values.append(item[column])
  datacolumns.append(column_header_and_values)

# Now that data columns has been filled with all remaining data, we now print all data to csv
# Initialize a variable that has the number of rows, should be same for all columns
num_rows = len(phone_numbers)

# Remove the space in the header Contact ID to give ContactID 
datacolumns[0] = [w.replace('Contact ', 'Contact') for w in datacolumns[0]]

# Initalize the file and a variable that contains all columns
resultFyle = open('out.csv','wb')
wr = csv.writer(resultFyle, dialect='excel')
all_columns = [phone_numbers] + datacolumns

# Print routine that writes to CSV  
for row_idx in range(num_rows):
  # The following three lines could be a list comprehension: row_to_write = [column[row_idx] for column in all_columns]
  row_to_write = []
  for column in all_columns:
    row_to_write.append(column[row_idx])
  
  wr.writerow(row_to_write)

# Convert csv to xlsx
merge_all_to_a_book(glob.glob("out.csv"), fileout)

print 'Your data has been massaged.'





