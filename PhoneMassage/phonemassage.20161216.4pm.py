# This script will take .xlsx (Excel) files and clean hyphens in the phone number entries
# and a variety of other unicode and unwanted characters. 
# This script will remove spaces in the headers and remove any OPT-IN to Text columns 
# after removing Opt-In = 'N' individuals from the data. 

# WHEN RUNNING THIS SCRIPT, PUT THE FOLLOWING COMMAND INTO THE COMMAND LINE

#                python phone.massage.py InputFile.csv OutPutFile.xlsx
#        where InputFile.xlsx is your input file, and OutPutFile.xlsx your ouput

# Version 1.2 - This version is compatible with Python3.*

import xlrd
import re
from xlrd import open_workbook, cellname
import csv
import string
import sys
import itertools
import openpyxl
import glob
import os
from pyexcel.cookbook import merge_all_to_a_book


# For Microsoft Users, run the following commands after you have installed the principle python library
# pip install xlrd
# pip install openpyxl
# pip install pyexcel
# pip install pyexcel-xlsx 

# This shortcut deals with xrange not available in python3

if sys.version_info > (3,):
    xrange = range

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
common = list(map(str, common)) 

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
	print ('Please Change the header of you phone numbers to Mobile Number')

# The following chunk cleans up the phone numbers

# Convert from unicode to strings - map needs to have list(map()) in Python 3 and higher
phone_numbers = list(map(str, phone_numbers))

# THIS AREA COULD BE REWRITTEN WITH A FUNCTION BUT FOR NOW, IT FUNCTIONS

# For loop runs through phone number list and replaces bad symbols in phone numbers
symbolstodelete = ['-', '(', ')', '#', '*', '?', '!', '@', '~', ' ']
for symbol in symbolstodelete:
	if symbol in phone_numbers:
		phone_numbers = phone_numbers.replace(symbol, '').replace(' ', '')

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


### UNICODE REMOVAL SECTION ###


# We need to remove the weird acute accent symbol in the 'you'll' message string. 
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

# There are names with an n with a ~ on top, we have to replace with an n to avoid unicode problems
unicodentilda = u'\u00F1'
for myitem in dict_list:
	for key in myitem:
		if key == 'Last Name' or key == 'First Name':
			myitem[key] = myitem[key].replace(unicodentilda, "n")

unicodeoum = u'\u00F6'
for myitem in dict_list:
	for key in myitem:
		if key == 'Last Name' or key == 'First Name':
			myitem[key] = myitem[key].replace(unicodeoum, "o")

unicodeaum = u'\u00E4'
for myitem in dict_list:
	for key in myitem:
		if key == 'Last Name' or key == 'First Name':
			myitem[key] = myitem[key].replace(unicodeaum, "a")

# Could make a list of all unicode symbols that have an e as the base character and define a function that does this a little more cleanly

unicodeeum = u'\u00EB'
for myitem in dict_list:
	for key in myitem:
		if key == 'Last Name' or key == 'First Name':
			myitem[key] = myitem[key].replace(unicodeeum, "e")

# Remove a a dictionary if the Opt-In to Text value is 'N'
dict_list = [element for element in dict_list if element.get('Opt-In to Text', '') != 'N']


### COMPILE REST OF DATA AND PRINT ### NO MORE DATA EDITING AFTER THIS POINT ####


# We need to insert the rest of the data into the csv while perserving order.
# Remove the phone number header from the keys list.

stringkeys = list(map(str, keys)) 
stringkeys.remove(common)

# Remove the Optin Column after all 'N' values have been removed. We don't want to print to final csv
if len(stringkeys) > 2:
	stringkeys.pop()  

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

# Remove all remaining spaces in all headers found exclusively in datacolumns
for idx in range(len(stringkeys)):
	datacolumns[idx][0]= datacolumns[idx][0].replace(' ','')

# Initalize the file and a variable that contains all columns
resultFyle = open('out.csv','w')
wr = csv.writer(resultFyle, dialect='excel')
all_columns = [phone_numbers] + datacolumns

# Writing routine that writes to CSV  
for row_idx in range(num_rows):
  # The following three lines could be a list comprehension: row_to_write = [column[row_idx] for column in all_columns]
  row_to_write = []
  for column in all_columns:
    row_to_write.append(column[row_idx])
  
  wr.writerow(row_to_write)

# Convert csv to xlsx
merge_all_to_a_book(glob.glob("out.csv"), fileout)

print ('Your data has been massaged.')





