# -*- coding: utf-8 -*-
"""
IHN rematches tool 1.00
Created on Mon Sep 27 19:03:49 2020
"""

###############################################################################
#                                                                             #
# SETUP - Identyfying files and accessing data                                #
#                                                                             #
###############################################################################


# required packages and technical settings related to .df warnings
import os
import sys
import pandas as pd
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
pd.options.mode.chained_assignment = None

# setting up working directory - script should be run by using command:
# python C:\PRZYKLAD_SCIEZKI\rematching_matrix.py
os.chdir(os.path.dirname(sys.argv[0]))

# we can define functions which will be used later to convert data:
# 1) function for turning all necessary numbers from floats to integers:
def integer_convert(i):
    try:
        return int(i)
    except (ValueError, TypeError):
        return i
# 2) function that will lookup keywords in POSDATA.YOUR_TITLE:
def title_check(ytitle, keyword):
    return f'{keyword}' in f'{ytitle}'

# following lines will ask user to enter the name of the excel file
EXCEL_NAME = input('\n   enter the name of excel file: ')
print('\n      importing posdata...')

# sheet Data contains incumbents information and will be our main framework
DATA = pd.read_excel(EXCEL_NAME, sheet_name='Data')
DATA.ORGDATA_TYPE_CARE.fillna(0, inplace=True)
DATA.ORGDATA_TYPE_CARE = [integer_convert(i) for i in DATA.ORGDATA_TYPE_CARE]

# keyword matrix contains map for rematching by YOUR_TITLE values:
print('      preparing Keyword matrix...')
KEYWORD_MATRIX = pd.read_excel(EXCEL_NAME, sheet_name='Keyword')
KEYWORD_MATRIX.fillna('unchanged', inplace=True)
KEYWORD_MATRIX.POS_CODE_DT = [integer_convert(i) for i in
                              KEYWORD_MATRIX.POS_CODE_DT]

# care type matrix contains map for rematching by incumbents' TYPE_CARE values:
print('      preparing Type of Care matrix...')
CARE_TYPE_MATRIX = pd.read_excel(EXCEL_NAME, sheet_name='CareType')
CARE_TYPE_MATRIX.fillna('unchanged', inplace=True)
CARE_TYPE_MATRIX.POS_CODE_DT = [integer_convert(i) for i in
                                CARE_TYPE_MATRIX.POS_CODE_DT]


###############################################################################
#                                                                             #
# Rematches by YOUR_TITLE                                                     #
#                                                                             #
###############################################################################


print('\n      matching by YOUR_TITLE:')

# we use pythonic vlookup to match incumbents with list of rematches possible
# for their POS_CODEs:
DATA = pd.merge(DATA, KEYWORD_MATRIX, on='POS_CODE_DT', how='left')

# we create a column which will show whether rematch has been done - by default
# it will be displaying lack of changes:
DATA['POS_CODE_KEYWORD'] = 'unchanged'

# now we can prepare a list of keywords, using names of columns from XLS file:
KEYWORDS_LIST = list(map(lambda i:i.upper(), KEYWORD_MATRIX.columns))
KEYWORDS_LIST.remove('POS_CODE_DT')

# we will check POSDATA by list of YTITLE fragments that may require rematching
# and create dictionaries allowing us to find new POS_CODES.
# In case of YOUR_TITLE containing multiple words from list, the LAST match
# will be used - we can also save only the FIRST one by using 93rd line of code
print('      creating dictionary...')

DATA_MERGER = pd.DataFrame()
for i in KEYWORDS_LIST:
    YTITLE = set([title for title in DATA.YOUR_TITLE if title_check(title, i)])
    LOOKUP = KEYWORD_MATRIX.set_index('POS_CODE_DT')[i].to_dict()
    DATA_FILTER = DATA[DATA.YOUR_TITLE.isin(list(YTITLE))]
#    DATA_FILTER = DATA_FILTER[DATA_FILTER.POS_CODE_KEYWORD == 'unchanged']
    if DATA_FILTER.OBJECTID.size != 0:
        DATA_FILTER['POS_CODE_KEYWORD'] = DATA.POS_CODE_DT.map(LOOKUP)
    DATA_MERGER = DATA_MERGER.append(DATA_FILTER)

DATA_MERGER = DATA_MERGER[DATA_MERGER.POS_CODE_KEYWORD != 'unchanged']
LOOKUP = DATA_MERGER.set_index('OBJECTID')['POS_CODE_KEYWORD'].to_dict()
DATA['POS_CODE_KEYWORD'] = DATA.OBJECTID.map(LOOKUP)
del DATA_MERGER, DATA_FILTER, LOOKUP


###############################################################################
#                                                                             #
# Rematches by Type of Care                                                   #
#                                                                             #
###############################################################################


print('\n      matching by Type of Care:')

# we use pythonic vlookup to match incumbents with list of rematches possible
# for their POS_CODEs:
DATA = pd.merge(DATA, CARE_TYPE_MATRIX, on='POS_CODE_DT', how='left')

# we create a column which will show whether rematch has been done - by default
# it will be displaying lack of changes:
DATA['POS_CODE_CARE'] = 'unchanged'

# then prepare a list of care type options, using column names from XLS file:
CARE_TYPE_LIST = list(CARE_TYPE_MATRIX.columns)
CARE_TYPE_LIST.remove('POS_CODE_DT')

# as care types are integers, we do not have to use separate function verifying
# parts of strings etc - instead we can lookup incumbents by their 
# ORGDATA_TYPE_CARE, using similar dictionaries as for keywords:
print('      creating dictionary...')

DATA_MERGER = pd.DataFrame()
for i in CARE_TYPE_LIST:
    LOOKUP = CARE_TYPE_MATRIX.set_index('POS_CODE_DT')[i].to_dict()
    DATA_FILTER = DATA[DATA.ORGDATA_TYPE_CARE == i]
    if DATA_FILTER.OBJECTID.size != 0:
        DATA_FILTER['POS_CODE_CARE'] = DATA.POS_CODE_DT.map(LOOKUP)
    DATA_MERGER = DATA_MERGER.append(DATA_FILTER)

DATA_MERGER = DATA_MERGER[DATA_MERGER.POS_CODE_CARE != 'unchanged']
LOOKUP = DATA_MERGER.set_index('OBJECTID')['POS_CODE_CARE'].to_dict()
DATA['POS_CODE_CARE'] = DATA.OBJECTID.map(LOOKUP)
del DATA_MERGER, DATA_FILTER, LOOKUP


###############################################################################
#                                                                             #
# Preparing final XLS for PR and CSVs for import                              #
#                                                                             #
###############################################################################


print('\n      preparing final excel files...')

# before saving the file we can reposition the columns for better readability:
KEYWORD_COL = DATA.pop('POS_CODE_KEYWORD')
CARE_COL = DATA.pop('POS_CODE_CARE')
DATA.insert(5, "POS_CODE_KEYWORD", KEYWORD_COL)
DATA.insert(6, "POS_CODE_CARE", CARE_COL)

del KEYWORD_COL, CARE_COL

# we add final conversions to ensure that all numeric POS_CODEs are integers:
DATA.POS_CODE_KEYWORD = [integer_convert(i) for i in DATA.POS_CODE_KEYWORD]
DATA.POS_CODE_CARE = [integer_convert(i) for i in DATA.POS_CODE_CARE]

# now we can save output to file allowing for quick review of data:
DATA.to_excel('whole_data_for_PR.xlsx', index=False)

# we can also prepare CSV files for quick import, if PR confirms data is ok:
DATA_K = DATA[['OBJECTID', 'POS_CODE_KEYWORD']].copy()
DATA_K = DATA_K[DATA_K['POS_CODE_KEYWORD'] != 'unchanged']
DATA_K.rename(columns={'OBJECTID': 'ObjID', 'POS_CODE_KEYWORD': 'POS_CODE_DT'},
              inplace=True)
DATA_K.to_csv('by_keywords_import.csv', index=False, encoding='utf-8')

DATA_C = DATA[['OBJECTID', 'POS_CODE_CARE']].copy()
DATA_C = DATA_C[DATA_C['POS_CODE_CARE'] != 'unchanged']
DATA_C.rename(columns={'OBJECTID': 'ObjID', 'POS_CODE_CARE': 'POS_CODE_DT'},
              inplace=True)
DATA_C.to_csv('by_care_type_import.csv', index=False, encoding='utf-8')

print('\n   done')
