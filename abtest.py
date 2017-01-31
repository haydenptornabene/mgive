# The following routine is an automated A/B Testing process. The end goal 
# is to create a tool that can be accessed from the mGive dashboard and conduct 
# specalized A/B testing.

# VERSION NUMBER: 1.0.0

# This version will be in house only and it requires the input of an .xlsx file. 
# All suppression criteria are built into this script; eventually we would like the 
# suppression criteria to be factored into the first menu of the tool once it's in the 
# dashboard. 

# Version 1.0.0 should only be used for Red Cross

# In general, this routine: 
# a1. Delete any A/B Testing Column in the Client Database 
#  1. Reads an xlsx file 
#  2. Suppresses duplicates, donors who have donated in the last 15 days, or are subscribed to Biomed
#  3. Ascrbies randomized A or B values to all members in list
#  4. Updates the Client Database with values and a new time stamp column (for filtering) 
#  5. Prints an updated Excel file for bookkeeping.     

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
import random


# For new users, run the following commands after you have installed the principle python library
# pip install xlrd
# pip install openpyxl
# pip install pyexcel
# pip install pyexcel-xlsx 

# unix code: python abtest.py mobilesubscribers.file mobiledonations.file biomed.file output.file

######## PREAMBLE ##########

# This will take the second to last argument in the command line, the input file name
subscribersin = sys.argv[-3]
donationsin = sys.argv[-2]
biomedin = sys.argv[-1]


# This will take the last argument in the command line, the output file name
#fileout = sys.argv[-1] 

# Open the desired file
booksub = open_workbook(subscribersin)
sheetsub = booksub.sheet_by_index(0)

# Open the desired file
bookusr = open_workbook(donationsin)
sheetusr = bookusr.sheet_by_index(0)

# Open the desired file
bookbio = open_workbook(biomedin)
sheetbio = bookbio.sheet_by_index(0)


#CREATE THE DICTIONARIES 
# Create a dictionary with every value from the spreadsheet from the subscriber xlsx
subkeys = [sheetsub.cell(0, col_index).value for col_index in xrange(sheetsub.ncols)]
sub_dict = []
for row_index in xrange(1, sheetsub.nrows):
    d = {subkeys[col_index]: sheetsub.cell(row_index, col_index).value 
         for col_index in xrange(sheetsub.ncols)}
    sub_dict.append(d)
# Create a dictionary with every value from the spreadsheet from the donations xlsx
usrkeys = [sheetusr.cell(0, col_index).value for col_index in xrange(sheetusr.ncols)]
usr_dict = []
for row_index in xrange(1, sheetusr.nrows):
    d = {usrkeys[col_index]: sheetusr.cell(row_index, col_index).value 
         for col_index in xrange(sheetusr.ncols)}
    usr_dict.append(d)
# Create a dictionary with every value from the spreadsheet from the biomed xlsx
biokeys = [sheetbio.cell(0, col_index).value for col_index in xrange(sheetbio.ncols)]
bio_dict = []
for row_index in xrange(1, sheetbio.nrows):
    d = {biokeys[col_index]: sheetbio.cell(row_index, col_index).value 
         for col_index in xrange(sheetbio.ncols)}
    bio_dict.append(d)

# REMOVE DUPLICAES 
# Now that all three text files have been uploaded, we can begin to remove duplicates.
# We want to remove: 
# 1. Duplicates 2. Donors who made donations in the last 15 days 3. Biomed Subscribers 
# The sub_dict is all people subscribed in some capacity to Red Cross, the master so to speak
# We want to remove anyone from sub_dict that appears on usr_dict or bio_dict 
# (Donated in past 15 days or is a biomed subscriber)

# Extract mobile numbers from each list
sub_mobiles = [d['Mobile Number'] for d in sub_dict]
usr_mobiles = [d['Mobile Number'] for d in usr_dict]
bio_mobiles = [d['MobileNumber'] for d in bio_dict]

# Create an intersection between sub_mobiles and usr_mobiles. 
# Create new list with intersections removed.
sub_minus_usr_mobiles = []
sub_usr_intersect = set(sub_mobiles).intersection(usr_mobiles)
for i in sub_mobiles:
	if i not in sub_usr_intersect:
		sub_minus_usr_mobiles.append(i)

# Create an intersection between sub_minus_usr_mobiles and bio_mobiles. 
# Create new  list with intersections removed. 
cleanned_mobiles = []
subusr_bio_intersect = set(sub_minus_usr_mobiles).intersection(bio_mobiles)
for i in sub_minus_usr_mobiles:
	if i not in subusr_bio_intersect:
		cleanned_mobiles.append(i)

# Remove duplicates from cleanned_mobiles
master_mobiles = []
for i in cleanned_mobiles:
	if i not in master_mobiles:
		master_mobiles.append(i)

#Initalize AB Randomizer Variables
half = len(master_mobiles)/2
oddhalf = half + 1  
A = 'A'
B = 'B'
alist = []
blist = []

# Create two lists, one of A's and one of B's that handles lists with an odd and even size length
if len(master_mobiles) % 2 == 1: # Odd Case
	alist = [A] * half
	blist = [B] * oddhalf
if len(master_mobiles) % 2 == 0: # Even Case
	alist = [A] * half
	blist = [B] * half
# Concatenate lists
ABlist = alist+blist
# Randomize list
random.shuffle(ABlist)

# We create a dictionary with mobile numbers as the key and A or B as the value.
cleanned_mobiles_dictionary = dict(zip(master_mobiles, ABlist))

# Print dictionary to csv file to be read into SQL database. 
with open('ABTestingOutput.csv', 'wb') as csv_file:
    writer = csv.writer(csv_file)
    for key, value in cleanned_mobiles_dictionary.items():
       writer.writerow([key, value])

print 'Your AB Test spreadsheet is ready.'









