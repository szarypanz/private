
################ SETUP ########################################################

import os
import pandas as pd

# pointing to directory where the CSV file is located
PATH = 'C:\\Users\\krzysztof-ibisz\\Desktop\\python\\NFM'
os.chdir(PATH)
del PATH

# loading database as frame - remember to change name as needed.
DATA = pd.read_csv('Duplicates_posdata.csv')
# please remember to export CSV in UTF-8 encoding or, alternatively, use xlsx:
# DUPLICATES_DF = pd.read_excel('duplikaty_us.xlsx')

# removing rows with empty INCNO & EEID:
DATA = DATA[~DATA['INCNO'].isnull()]
DATA = DATA[~DATA['YOUR_EEID'].isnull()]

# using list of surveys we can determine their priority, using numbers;
# in case of duplicate YOUR_EEID_1 and YOUR_EEID_2 the BIGGER number will be
# marked as duplicate
DATA['PRIORITY'] = DATA['SURVEY_CODE']
DATA['PRIORITY'].replace(['srv1', 'srvb', 'srvx', 'sryz'],
                         [4, 3, 2, 1], inplace=True)


################ FINDING EEID DUPLICATES ######################################


# first we mark duplicated EEIDs in a new column:
DATA['EEID_d'] = DATA.YOUR_EEID.duplicated(keep=False)

# then we create subset containing only duplicated incumbents
EEID_DF = DATA[DATA['EEID_d']]
EEID_DF = EEID_DF.sort_values(by=['YOUR_EEID', 'CO_CODE1', 'PRIORITY'])

# adding keys which will be used to identify duplicates per specific Companies:
UNIQUE_KEY = ['YOUR_EEID', 'CO_CODE1']
EEID_DF['key'] = EEID_DF[UNIQUE_KEY].apply(
        lambda row: '_'.join(row.values.astype(str)), axis=1)

# we can now use function to identify correct values by PRIORITY column and
# create a separate frame containing ONLY duplicates which will receive flag:
EEID_CORRECT = EEID_DF.drop_duplicates(subset='key', keep='first')
EEID_DUPLICATE = EEID_DF[(~EEID_DF.OBJECTID.isin(EEID_CORRECT.OBJECTID))]
EEID_DUPLICATE['DUPLICATE_FLAG'] = 1

# now we can get leave only columns required for import into GST:
#EEID_DUPLICATE = EEID_DUPLICATE.filter(['OBJECTID', 'DUPLICATE_FLAG'], axis=1)

# we save results which will be available in the same folder:
EEID_DUPLICATE.to_csv('Results_EEID.csv', index=False, encoding='utf-8')


################ FINDING INCNO DUPLICATES #####################################


# first we mark duplicated INCNO in a new column:
DATA['INCNO_d'] = DATA.INCNO.duplicated(keep=False)

# then we create subset containing only duplicated incumbents
INCNO_DF = DATA[DATA['INCNO_d']]
INCNO_DF = INCNO_DF.sort_values(by=['INCNO', 'CO_CODE1', 'PRIORITY'])

# adding keys which will be used to identify duplicates per specific Companies:
UNIQUE_KEY = ['INCNO', 'CO_CODE1']
INCNO_DF['key'] = INCNO_DF[UNIQUE_KEY].apply(
        lambda row: '_'.join(row.values.astype(str)), axis=1)

# we can now use function to drop duplicates and specify (thanks to PRIORITY)
# which records will receive a duplicate flag:
INCNO_CORRECT = INCNO_DF.drop_duplicates(subset='key', keep='first')
INCNO_DUPLICATE = INCNO_DF[(~INCNO_DF.OBJECTID.isin(INCNO_CORRECT.OBJECTID))]
INCNO_DUPLICATE['DUPLICATE_FLAG'] = 1

# now we can get leave only columns required for import into GST:
#INCNO_DUPLICATE = INCNO_DUPLICATE.filter(['OBJECTID', 'DUPLICATE_FLAG'],
#                                         axis=1)

# we save results which will be available in the same folder:
INCNO_DUPLICATE.to_csv('Results_INCNO.csv', index=False, encoding='utf-8')
