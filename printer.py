
##############################################################################
#        ____  ______   ______            __                                 #
#       / __ \/ ____/  /_  __/___  ____  / /                                 #
#      / /_/ / /        / / / __ \/ __ \/ /                                  #
#     / ____/ /___     / / / /_/ / /_/ / /                                   #
#    /_/    \____/    /_/  \____/\____/_/   version 0.99                     #
#                                                                            #
##############################################################################

import os
import pandas as pd
import openpyxl as pyxl
import datetime
import sys

##############################################################################
#                                                                            #
# Preparations:                                                              #
#   → creating objects                                                       #
#   → loading settings file                                                  #
#   → specifying paths, working directories etc                              #
#                                                                            #
##############################################################################

# script can be initialized by typing:
# python H:\EXAMPLE\EXAMPLE\EXAMPLE script.py
# in anaconda prompt; make sure folder in that directory contains
# settings file and all exported GST queries

# setting working directory to the location of .py script file
os.chdir(os.path.dirname(sys.argv[0]))

Setup = pd.read_excel('Settings.xlsm')  # file has to be in the same loc as script
Path = os.path.dirname(sys.argv[0])     # same for XLS files
#Path = os.getcwd()    # commented out: alt Path for convenient testing  in IDE
SurveyFullName = Setup.iloc[9,2]
TemplateName = Setup.iloc[10,2]
SubmissionDeadline = Setup.iloc[11,2].date()
InternalDeadline = Setup.iloc[12,2].date()
QAD = Setup.iloc[13,2].date()

ExcEntities_YN = Setup.iloc[17,2]
ExcIncumbents_YN = Setup.iloc[18,2]
AGCRematches_YN = Setup.iloc[19,2]
MultipleYTPos_YN = Setup.iloc[20,2]
MultipleYTAll_YN = Setup.iloc[21,2]
Deadlines_YN = Setup.iloc[22,2]
QAD_YN = Setup.iloc[23,2]

FileNames = Setup.iloc[26:34,2].tolist()
ReportName = 'Problematic Companies {} - report.xlsx'.format(SurveyFullName)
TemplatePath = os.path.join(Path, TemplateName)

Summary = []
ExcEntitiesData = []
ExcReasonData = []
ExcIncumbentsData = []
AGCRematchesData = []
MultipleYTPositionsData = []
MultipleYTPositionsAll = []
AfterDeadlinesData = []
MDAData = []
SummaryIncumbents = []
LateContact = []
LateSub = []

tresholds = Setup.iloc[36:44, 2:4]


# checking if any tresholds should be skipped:
def create_tresholds():



    # duplicates tresholds is 'less than' condition so we'd replace empty/N
    # with 100%:
    if (tresholds.iloc[4,1] == 'N' or pd.isna(tresholds.iloc[4,1])):
        tresholds.iloc[4,1] = 1

    # custom default tresholds for YTPos & QAD
    if (tresholds.iloc[7,1] == 'N' or pd.isna(tresholds.iloc[7,1])):
        tresholds.iloc[7,1] = 14    # 14 days after QAD

    if (tresholds.iloc[6,1] == 'N' or pd.isna(tresholds.iloc[6,1])):
        tresholds.iloc[6,1] = 10    # 10 unique Your Titles

    # other N/missing tresholds can be replaced with 0:
    tresholds.replace('N', 0, inplace=True)
    tresholds.fillna(value=0, inplace=True)




# function responsible for loading data from exported GST queries
def import_xlsx_files(Path):

    print('getting data from excel files...')

    global Summary
    Summary = pd.read_excel(os.path.join(Path, FileNames[0]))

    global ExcEntitiesData
    if ExcEntities_YN == 'Y' and pd.notna(FileNames[1]):
        ExcEntitiesData = pd.read_excel(os.path.join(Path, FileNames[1]))

    global ExcReasonData
    if ExcEntities_YN == 'Y' and pd.notna(FileNames[2]):
        ExcReasonData = pd.read_excel(os.path.join(Path, FileNames[2]))

    global ExcIncumbentsData
    if ExcIncumbents_YN == 'Y' and pd.notna(FileNames[3]):
        ExcIncumbentsData = pd.read_excel(os.path.join(Path, FileNames[3]))

    global AGCRematchesData
    if AGCRematches_YN == 'Y' and pd.notna(FileNames[4]):
        AGCRematchesData = pd.read_excel(os.path.join(Path, FileNames[4]))

    global MultipleYTPositionsAll
    if (MultipleYTAll_YN == 'Y' or MultipleYTPos_YN == 'Y') and pd.notna(FileNames[5]):
        MultipleYTPositionsAll = pd.read_excel(os.path.join(Path, FileNames[5]))

    global LateSub
    global LateContact
    if Deadlines_YN == 'Y' and pd.notna(FileNames[6]) and pd.notna(FileNames[7]):
        LateContact = pd.read_excel(os.path.join(Path, FileNames[6]))
        LateSub = pd.read_excel(os.path.join(Path, FileNames[7]))

    print('transforming data...')




##############################################################################
#                                                                            #
# Main body                                                                  #
#   → data transformations                                                   #
#   → final XLSX files creation                                              #
#                                                                            #
##############################################################################


# making sure columns are named correctly - e.g for surveys that are using
# CPY_CODE instead of CO_CODE2 etc etc etc
def make_names_standard():

    global Summary
    SummaryColumns = Setup.iloc[0:12]['SummaryColumns']
    Summary.columns = SummaryColumns.dropna()

    if ExcEntities_YN == 'Y':
        global ExcEntitiesData
        ExcEntitiesColumns = Setup.iloc[0:6]['ExcEntitiesColumns']
        ExcEntitiesData.columns = ExcEntitiesColumns.dropna()

        global ExcReasonData
        ExcEntitiesReasonColumns = Setup.iloc[0:7]['ExcEntitiesReasonColumns']
        ExcReasonData.columns = ExcEntitiesReasonColumns.dropna()

    if ExcIncumbents_YN == 'Y':
        global ExcIncumbentsData
        ExcIncumbentsColumns = Setup.iloc[0:38]['ExcIncumbentsColumns']
        ExcIncumbentsData.columns = ExcIncumbentsColumns.dropna()

    if AGCRematches_YN == 'Y':
        global AGCRematchesData
        AGCRematchesColumns = Setup.iloc[0:5]['AGCRematchesColumns']
        AGCRematchesData.columns = AGCRematchesColumns.dropna()

    if MultipleYTAll_YN == 'Y' or MultipleYTPos_YN == 'Y':
        global MultipleYTPositionsAll
        MultipleYTAllColumns = Setup.iloc[0:10]['MultipleYTAllColumns']
        MultipleYTPositionsAll.columns = MultipleYTAllColumns.dropna()




# Excluded Entities sheet - adding perc of excluded entities
# and filtering out records below the treshold:
def create_excluded_entities():

    if ExcEntities_YN == 'Y':

        global ExcEntitiesData
        ExcEntitiesData = ExcEntitiesData[ExcEntitiesData.SUBMITTED
                                          >= tresholds.iloc[1,1]]

        ExcEntitiesData.loc[:,'EXCLUDED_PERC'] = round(
            ExcEntitiesData['EXCLUDED']
            / ExcEntitiesData['SUBMITTED'], 2)

        ExcEntitiesData = ExcEntitiesData[ExcEntitiesData.EXCLUDED_PERC
                                          >= tresholds.iloc[0,1]]




# preparing clean Exclude reasons (by co_code2), grouping them by SubID,
# matching them with Excluded Entities 
def create_exclude_reasons():

    global ExcReasonData

    if ExcEntities_YN == 'Y' and len(ExcReasonData) > 0:


        ExcReasonData['EXCLUDE_REASON'] = ExcReasonData['EXCLUDE_REASON'].fillna(value='Unspecified exclude reason')
        ExcReasonData['CO_NAME2'] = ExcReasonData['CO_NAME2'].fillna(value='Unnamed Entity')
        ExcReasonData['EXCLUDE_REASON'] = (ExcReasonData['CO_NAME2']+": "
                                           +ExcReasonData['EXCLUDE_REASON'])

        global ExcEntitiesData
        for x in ExcReasonData['SUBMISSIONID'].unique():
            data = ExcReasonData.loc[ExcReasonData['SUBMISSIONID'] == x]
            sep = (';'+'\n')
            z = sep.join(data['EXCLUDE_REASON'])
            ExcEntitiesData.loc[ExcEntitiesData['CONTACT_SUBMISSIONID'] == x,
                                'EXCLUDE_REASON'] = z
            ExcEntitiesData['EXCLUDE_REASON'].fillna(value='', inplace=True)

        # removing Reasons dataframe to free some memory
        del ExcReasonData




# calculating Excluded Incumbents data & total incumbents for Summary tab,
# applying tresholds to data
def create_excluded_incumbents():

    if ExcIncumbents_YN == 'Y':

        global ExcIncumbentsData
        global SummaryIncumbents

        # creating small df with info about total incumbents for Summary tab
        SummaryIncumbents = ExcIncumbentsData.loc[:,['ORGDATA_SUBMISSIONID','ALLINCUMBENTS']]
        SummaryIncumbents = SummaryIncumbents[~SummaryIncumbents['ORGDATA_SUBMISSIONID'].isnull()]
        SummaryIncumbents = SummaryIncumbents.groupby(['ORGDATA_SUBMISSIONID']).sum()
        SummaryIncumbents['SUBID'] = SummaryIncumbents.index

        ExcIncumbentsData.loc[:,'EXCLUDED_PERC'] = round(
                ExcIncumbentsData['EXCLUDED_INCUMBENTS']
                / ExcIncumbentsData['ALLINCUMBENTS'], 2)

        # checking treshold for duplicates requires creating temporary column
        # for calculations
        y = 0
        for x in ExcIncumbentsData.EF_2_COUNT:
            if x > 0:
                ExcIncumbentsData.loc[y,"dupl_perc"] = (ExcIncumbentsData.iloc[y,12]
                                                        / sum(ExcIncumbentsData.iloc[y,9:38]))
            else:
                ExcIncumbentsData.loc[y,"dupl_perc"] = 0
            y = y + 1

        ExcIncumbentsData = ExcIncumbentsData[ExcIncumbentsData.ALLINCUMBENTS
                                              >= tresholds.iloc[3,1]]
        ExcIncumbentsData = ExcIncumbentsData[ExcIncumbentsData.EXCLUDED_PERC
                                              >= tresholds.iloc[2,1]]
        ExcIncumbentsData = ExcIncumbentsData[ExcIncumbentsData.dupl_perc
                                              <= tresholds.iloc[4,1]]
        ExcIncumbentsData.drop(columns='dupl_perc', inplace=True)




# YTPositions tab requires only tresholds application
def create_YTPos():

    if MultipleYTPos_YN == 'Y':

        global MultipleYTPositionsData
        global MultipleYTPositionsAll

        # transforming raw YTAll data:
        MultipleYTPositionsData = MultipleYTPositionsAll.copy()
        MultipleYTPositionsData.drop(columns='ALLINCUMBENTS', inplace=True)
        MultipleYTPositionsData = MultipleYTPositionsData.groupby([
                'ORGDATA_SUBMISSIONID','ORGDATA_CO_CODE1','ORGDATA_CO_CODE2',
                'ORGDATA_CO_NAME2','POS_CODE','POSTITLE']).sum()
        MultipleYTPositionsData = MultipleYTPositionsData.reset_index()

        # applying YTitles treshold
        MultipleYTPositionsData = MultipleYTPositionsData[MultipleYTPositionsData.UNIQUEYOURTITLE
                                                          >= tresholds.iloc[6,1]]

        # creating temporary object to apply 'x Positions per Company' treshold
        Cpys = MultipleYTPositionsData.ORGDATA_CO_CODE2.value_counts().to_frame()
        Cpys = Cpys.ORGDATA_CO_CODE2[Cpys.ORGDATA_CO_CODE2
                                     >= tresholds.iloc[5,1]].to_frame()
        Cpys['key'] = Cpys.index

        # filtering final DF by companys that are fullfiling Pos/Cpy treshold
        MultipleYTPositionsData = MultipleYTPositionsData[
                MultipleYTPositionsData['ORGDATA_CO_CODE2'].isin(Cpys['key'])]


        # if we do not need the other YT tab, we could remove the source df
        # to free some memory:
        if MultipleYTAll_YN != 'Y':
            del MultipleYTPositionsAll




def create_YTAll():

    if MultipleYTAll_YN == 'Y':

        global MultipleYTPositionsAll

        # applying upper case, cleaning yt from illegal characters
        MultipleYTPositionsAll['YOUR_TITLE'] = MultipleYTPositionsAll['YOUR_TITLE'].str.upper()
        MultipleYTPositionsAll = MultipleYTPositionsAll.applymap(
                lambda x: x.encode('unicode_escape').decode('utf-8') if isinstance(x, str) else x)

        # filtering out records that are not present on YTPos, if it is available:
        if MultipleYTPos_YN == 'Y':

            MultipleYTPositionsAll = MultipleYTPositionsAll[
                MultipleYTPositionsAll.ORGDATA_CO_CODE2.isin(MultipleYTPositionsData.ORGDATA_CO_CODE2)]

        # checking if given YTitle is problematic (i.e. appears on YT Pos Data)
        MultipleYTPositionsAll['PROBLEMATIC'] = 'N'
        
        if MultipleYTPos_YN == 'Y':

            # we want to check whether specific poscodes from *specific*
            # companies are present on YT Pos Data tab
            MultipleYTPositionsAll['key'] = (MultipleYTPositionsAll['ORGDATA_CO_CODE1'].map(str)
                                             + MultipleYTPositionsAll['POS_CODE'])
            
            MultipleYTPositionsData['key'] = (MultipleYTPositionsData['ORGDATA_CO_CODE1'].map(str)
                                              + MultipleYTPositionsData['POS_CODE'])

            check = MultipleYTPositionsData.key.unique()

            for x in MultipleYTPositionsAll['key'].unique():
                if x in check:
                    MultipleYTPositionsAll.loc[MultipleYTPositionsAll['key']
                                               == x,'PROBLEMATIC'] = 'Y'

            MultipleYTPositionsData.drop(columns='key', inplace=True)
            MultipleYTPositionsAll.drop(columns='key', inplace=True)




# Creating Data for 'Deadlines' tab - downloading CONTACT info and calculating
# degree and type of delays.
# As a first step we need to have function for date conversion
# (GST exports store date in excel-like timestamp format)
def date_convert(gst_date):

    return(datetime.datetime(1899, 12, 30) + datetime.timedelta(days=gst_date))


# Then we can use it in creating After Deadline Tab
def create_after_deadline():

    if Deadlines_YN == 'Y':

        global AfterDeadlinesData
        AfterDeadlinesData = Summary.loc[:,['SUBMISSIONID', 'CO_CODE1',
                                            'CONTACTNAME', 'CO_NAME1',
                                            'CN_TITLE', 'CN_PHONE',
                                            'CONTACTEMAIL', 'IMPORTDATE',]]

    # converting submission date into more readable format
        for x in AfterDeadlinesData['IMPORTDATE']:
            AfterDeadlinesData.loc[AfterDeadlinesData['IMPORTDATE'] == x,
                                   'IMPORTDATE'] = date_convert(x).date()

    # calculating if submission breached either of 2 deadlines
        AfterDeadlinesData['checkCLIENT'] = AfterDeadlinesData['IMPORTDATE'] - SubmissionDeadline
        AfterDeadlinesData['checkCLIENT'] = AfterDeadlinesData['checkCLIENT'].dt.days
        AfterDeadlinesData['checkINT'] = AfterDeadlinesData['IMPORTDATE'] - InternalDeadline
        AfterDeadlinesData['checkINT'] = AfterDeadlinesData['checkINT'].dt.days

    # creating columns in which we will store data about the type of delay
        AfterDeadlinesData['DEADLINE_BREACH'] = ''
        AfterDeadlinesData['DELAY'] = ''

    # creating loop which will determine what type of delay should be
    # picked for each sub
        y = 0
        for x in AfterDeadlinesData['SUBMISSIONID']:
            client = AfterDeadlinesData.iloc[y,8]
            internal = AfterDeadlinesData.iloc[y,9]
            if internal > 0:
                AfterDeadlinesData.iloc[y,10] = 'Internal'
                AfterDeadlinesData.iloc[y,11] = internal
            elif client > 0:
                AfterDeadlinesData.iloc[y,10] = 'Client'
                AfterDeadlinesData.iloc[y,11] = client
            else:
                AfterDeadlinesData.iloc[y,10] = "REMOVE ROW"
            y = y + 1

    # removing temporary technical columns, filtering out correct records,
    # setting data in correct order
        AfterDeadlinesData.drop(columns=['checkINT','checkCLIENT'], inplace=True)
        AfterDeadlinesData = (
                AfterDeadlinesData[AfterDeadlinesData.DEADLINE_BREACH != "REMOVE ROW"])
        ColumnOrder = Setup.iloc[0:10]['AfterDeadlineOrder']
        AfterDeadlinesData = AfterDeadlinesData[ColumnOrder]




# Creating data for 'DVF after QAD' tab
def create_mda():

    if QAD_YN == 'Y':

        global MDAData

        MDAData = Summary.loc[:,['SUBMISSIONID', 'CO_CODE1', 'MERCER_ACTION',
                                 'CONTACTNAME','CO_NAME1','CN_TITLE',
                                 'CN_PHONE','CONTACTEMAIL','RETURNED_SAW',]]

        # Removing rows with no date (some records in MBD had those)
        MDAData = MDAData[~MDAData['RETURNED_SAW'].isnull()]

        # converting DVF date into more readable format
        for x in MDAData['RETURNED_SAW']:
            MDAData.loc[MDAData['RETURNED_SAW'] == x,
                        'RETURNED_SAW'] = date_convert(x).date()

        # calculating if DVF broke deadline
        MDAData['DELAY'] = MDAData['RETURNED_SAW'] - QAD
        MDAData['DELAY'] = MDAData['DELAY'].dt.days

        # pasting 0s for DVFs that returned before QAD and would have '-x'
        # results in 'DELAY' field. Replace 2nd '0' here with any string 
        # or message you want to display:
        MDAData.loc[MDAData['DELAY'] < 0, 'DELAY'] = 0


        # preparing QAD Y/N column
        MDAData['QAD_BREACH'] = ''
        y = 0
        for x in MDAData['SUBMISSIONID']:
            qad_check = MDAData.iloc[y,9]
            if qad_check > int(tresholds.iloc[7,1]):  # cutoff as per Treshold
                MDAData.iloc[y,10] = "Y"
            else:
                MDAData.iloc[y,10] = "N"
            y = y + 1

        # Preparing MDA Y/N column
        MDAData.loc[~MDAData.MERCER_ACTION.isnull(), 'MERCER_ACTION'] = 'Y'
        MDAData['MERCER_ACTION'].fillna(value='N', inplace=True)

        # filtering out records which are neither afterQAD / MDA,
        # sorting, removing technical columns etc
        MDAData['remove_col'] = MDAData['MERCER_ACTION'] + MDAData['QAD_BREACH']
        MDAData = MDAData[MDAData.remove_col != "NN"]
        MDAData.drop(columns='remove_col', inplace=True)
        ColumnOrder = Setup.iloc[0:11]['MDAOrder']
        MDAData = MDAData[ColumnOrder]




# INSERTING CONTACT DATA
# Function that'll identify which sheets need to be updated with Contact data
# then paste it and rearrange columns in correct order
def create_contact_details():

    ContactSlice = Summary.loc[:,['SUBMISSIONID', 'CONTACTNAME',
                                  'CN_TITLE','CN_PHONE','CONTACTEMAIL']]

    if ExcEntities_YN == 'Y':
        global ExcEntitiesData
        ExcEntitiesData = ExcEntitiesData.merge(ContactSlice,
                                                left_on='CONTACT_SUBMISSIONID',
                                                right_on='SUBMISSIONID')
        ExcEntitiesData.drop(columns='SUBMISSIONID', inplace=True)
        ColumnOrder = Setup.iloc[0:12]['ExcEntitiesOrder']
        ExcEntitiesData = ExcEntitiesData[ColumnOrder]

    if ExcIncumbents_YN == 'Y':
        global ExcIncumbentsData
        ExcIncumbentsData = ExcIncumbentsData.merge(ContactSlice,
                                                    left_on='ORGDATA_SUBMISSIONID',
                                                    right_on='SUBMISSIONID')
        ExcIncumbentsData.drop(columns='SUBMISSIONID', inplace=True)
        ColumnOrder = Setup.iloc[0:43]['ExcIncumbentsOrder']
        ExcIncumbentsData = ExcIncumbentsData[ColumnOrder]

    if AGCRematches_YN == 'Y':
        global AGCRematchesData
        AGCRematchesData = AGCRematchesData.merge(ContactSlice,
                                                  left_on='SOURCE_ORGDATA_SUBMISSIONID',
                                                  right_on='SUBMISSIONID')
        AGCRematchesData.drop(columns='SUBMISSIONID', inplace=True)
        ColumnOrder = Setup.iloc[0:9]['AGCRematchesOrder']
        AGCRematchesData = AGCRematchesData[ColumnOrder]

    if MultipleYTPos_YN == 'Y':
        global MultipleYTPositionsData
        MultipleYTPositionsData = MultipleYTPositionsData.merge(ContactSlice,
                                                                left_on='ORGDATA_SUBMISSIONID',
                                                                right_on='SUBMISSIONID')
        ColumnOrder = Setup.iloc[0:12]['MultipleYTPosOrder']
        MultipleYTPositionsData = MultipleYTPositionsData[ColumnOrder]

    if MultipleYTAll_YN == 'Y':
        global MultipleYTPositionsAll
        MultipleYTPositionsAll = MultipleYTPositionsAll.merge(ContactSlice,
                                                              left_on='ORGDATA_SUBMISSIONID',
                                                              right_on='SUBMISSIONID')
        ColumnOrder = Setup.iloc[0:13]['MultipleYTAllOrder']
        MultipleYTPositionsAll = MultipleYTPositionsAll[ColumnOrder]




# when all data tabs are prepared, we can obtain necessary info for SUMMARY tab
def create_summary():

    # getting Excluded Entities stats
    global Summary
    if ExcEntities_YN == 'Y':
        ExcEntitiesSlice = ExcEntitiesData.loc[:,['CONTACT_SUBMISSIONID',
                                                  'EXCLUDED']]
        Summary = Summary.merge(ExcEntitiesSlice, left_on='SUBMISSIONID',
                                right_on='CONTACT_SUBMISSIONID',
                                how='left')
        Summary.drop(columns='CONTACT_SUBMISSIONID', inplace=True)
        Summary.EXCLUDED.fillna(value='n/a', inplace=True)
    else:
        Summary.loc[:,'EXCLUDED'] = 'n/a'


    # ...Excluded Incumbents...
    if ExcIncumbents_YN == 'Y':
        # as ExcInc data is per Co_code2 to prepare Summary we have
        # to roll it by SubmissionID
        ExIncSlice = ExcIncumbentsData.loc[:,['ORGDATA_SUBMISSIONID',
                                                      'EXCLUDED_INCUMBENTS']]
        ExIncSlice = ExIncSlice[~ExIncSlice['ORGDATA_SUBMISSIONID'].isnull()]
        ExIncSlice = ExIncSlice.groupby(['ORGDATA_SUBMISSIONID']).sum()
        ExIncSlice['SUBID'] = ExIncSlice.index

        Summary = Summary.merge(ExIncSlice, left_on='SUBMISSIONID',
                                right_on='SUBID',
                                how='left')
        Summary.drop(columns='SUBID', inplace=True)
        Summary.EXCLUDED_INCUMBENTS.fillna(value='n/a', inplace=True)

        Summary = Summary.merge(SummaryIncumbents, left_on='SUBMISSIONID',
                                right_on='SUBID',
                                how='left')
        Summary.drop(columns='SUBID', inplace=True)
        Summary.ALLINCUMBENTS.fillna(value='n/a', inplace=True)
    else:
        Summary.loc[:,'ALLINCUMBENTS'] = 'n/a'
        Summary.loc[:,'EXCLUDED_INCUMBENTS'] = 'n/a'


    # ...AGCRematches...
    if AGCRematches_YN == 'Y':
        # as AGCRematches data is per Co_code2 to prepare Summary we have
        # to roll it by SubmissionID
        AGCSlice = AGCRematchesData.loc[:,['SOURCE_ORGDATA_SUBMISSIONID',
                                        'OBJECTID_COUNT']]
        AGCSlice = AGCSlice[~AGCSlice['SOURCE_ORGDATA_SUBMISSIONID'].isnull()]
        AGCSlice = AGCSlice.groupby(['SOURCE_ORGDATA_SUBMISSIONID']).sum()
        AGCSlice['SUBID'] = AGCSlice.index

        Summary = Summary.merge(AGCSlice, left_on='SUBMISSIONID',
                                right_on='SUBID',
                                how='left')
        Summary.drop(columns='SUBID', inplace=True)
        Summary.OBJECTID_COUNT.fillna(value='n/a', inplace=True)
    else:
        Summary.loc[:,'OBJECTID_COUNT'] = 'n/a'

    # ...YTitle data...
    if MultipleYTPos_YN == 'Y':
        YTSlice = MultipleYTPositionsData.loc[:,['ORGDATA_SUBMISSIONID',
                                              'UNIQUEYOURTITLE']]
        # in Ytitles stats are by YOUR_TITLE etc, with duplicated SubIDs,
        # so we have to prepare a list of unique values
        z = YTSlice.iloc[:,0].value_counts().to_frame()
        z['SUBMISSIONID'] = z.index
        Summary = Summary.merge(z, left_on='SUBMISSIONID',
                                right_on='SUBMISSIONID',
                                how='left')
        Summary = Summary.rename(columns=
                                 {'ORGDATA_SUBMISSIONID':'UNIQUEYOURTITLE'})
        Summary.UNIQUEYOURTITLE.fillna(value='n/a', inplace=True)
    else:
        Summary.loc[:,'UNIQUEYOURTITLE'] = 'n/a'

    # ...Deadlines...
    if Deadlines_YN == 'Y':
        DeadlinesSlice = AfterDeadlinesData.loc[:,['SUBMISSIONID',
                                                'DEADLINE_BREACH']]
        Summary = Summary.merge(DeadlinesSlice, left_on='SUBMISSIONID',
                                right_on='SUBMISSIONID',
                                how='left')
        Summary.DEADLINE_BREACH.fillna(value='n/a', inplace=True)
    else:
        Summary.loc[:,'DEADLINE_BREACH'] = 'n/a'

    # ...QAD...
    if QAD_YN == 'Y':
        Summary.drop(columns='MERCER_ACTION', inplace=True)
        QADSlice = MDAData.loc[:,['SUBMISSIONID', 'DELAY', 'MERCER_ACTION',
                                  'QAD_BREACH']]

        y = 0                                            #
        for x in QADSlice.QAD_BREACH:                    # loop that'll block
            if x == 'N':                                 # submissions on time
                QADSlice.iloc[y,1] = 'n/a'               # from appearing
            else:                                        # on Summary
                pass                                     #
            y = y + 1                                    #
        QADSlice.drop(columns='QAD_BREACH', inplace=True)

        Summary = Summary.merge(QADSlice, left_on='SUBMISSIONID',
                                right_on='SUBMISSIONID',
                                how='left')
        Summary.DELAY.fillna(value='n/a', inplace=True)
        Summary.MERCER_ACTION.fillna(value='n/a', inplace=True)
        Summary.MERCER_ACTION.replace('N', 'n/a', inplace=True)
    else:
        Summary.loc[:,'DELAY'] = 'n/a'
        Summary.loc[:,'MERCER_ACTION'] = 'n/a'

    # marking  records available at this stage as *not* late submissions
    Summary.loc[:,'LATESUB'] = 'n/a'


    # removing records which, after applying tresholds, are not displaying any
    # issues and/or display them only on irrelevant tabs (e.g. QAD_YN = N, etc)
    ColumnOrder = Setup.iloc[0:16]['SummaryTempOrder']
    Summary = Summary[ColumnOrder]
    y = 0
    for x in Summary.SUBMISSIONID:
        z = Summary.iloc[y,8:14].unique()
        Summary.loc[y,'removal'] = str(z)
        y = y + 1
    Summary = Summary[Summary.removal != "['n/a']"]
    Summary.drop(columns='removal', inplace=True)


    # Setting correct column order and making sure their names are OK
    ColumnOrder = Setup.iloc[0:16]['SummaryOrder']
    Summary = Summary[ColumnOrder]




# Creating Late Sub data
def create_latesub():

    if Deadlines_YN == 'Y' and pd.notna(FileNames[6]) and pd.notna(FileNames[7]):

        global LateSub
        # limiting df only to late subs, dropping unnecessary columns
        LateSub = LateSub[LateSub.State == 'Late Submission']
        LateSub = LateSub.loc[:,['State','Submission Id']]
        LateSub = LateSub.rename(columns={'State':'LATESUB'})


        global LateContact
        LateContact = LateContact.merge(LateSub, left_on='SUBMISSIONID',
                                        right_on='Submission Id',
                                        how='left')

        LateContact = LateContact[LateContact['LATESUB'] == 'Late Submission']
        LateContact = LateContact.rename(columns={'SMIGROUPCODE':'CO_CODE1'})
        del LateSub     # removing redundant df


        # converting submission date into more readable format
        for x in LateContact['IMPORTDATE']:
            LateContact.loc[LateContact['IMPORTDATE'] == x,
                            'IMPORTDATE'] = date_convert(x).date()


        # calculating degree of delay
        LateContact['LATESUB'] = LateContact['IMPORTDATE'] - InternalDeadline
        LateContact['LATESUB'] = LateContact['LATESUB'].dt.days


        # transforming and pasting data into QAD/MDA tab:
        LateContactSlice = LateContact.copy()
        LateContactSlice = LateContactSlice.rename(columns={'LATESUB':'DELAY'})
        LateContactSlice['DEADLINE_BREACH'] = 'Late Submission'
        ColumnOrder = Setup.iloc[0:10]['AfterDeadlineOrder']
        LateContactSlice = LateContactSlice[ColumnOrder]
        global AfterDeadlinesData
        AfterDeadlinesData = AfterDeadlinesData.append(LateContactSlice,
                                                       sort=False)


        # next step: organizing data for Summary tab and pasting it there:
        LateContact.drop(columns='IMPORTDATE', inplace=True)
        LateContact['ALLINCUMBENTS'] = 'n/a'
        LateContact['EXCLUDED'] = 'n/a'
        LateContact['EXCLUDED_INCUMBENTS'] = 'n/a'
        LateContact['OBJECTID_COUNT'] = 'n/a'
        LateContact['UNIQUEYOURTITLE'] = 'n/a'
        LateContact['DEADLINE_BREACH'] = 'Late Submission'
        LateContact['DELAY'] = 'n/a'
        LateContact['MERCER_ACTION'] = 'n/a'

        ColumnOrder = Setup.iloc[0:16]['SummaryOrder']
        LateContact = LateContact[ColumnOrder]

        global Summary
        Summary = Summary.append(LateContact, sort=False)

        del LateContact     # removing redundant df




#   CREATING XLSX FILE FOR FINAL REPORT
#   function responsible for opening up Template and filling it with data
def create_report_file():

    print('opening template...')
    template = pyxl.load_workbook(filename=TemplateName)

    # Copying misc data (dates, names) from Settings to Introduction tab
    print('updating template...')
    sheet = template['Introduction']
    sheet['B1'].value = '2020 {} - Problematic Companies'.format(SurveyFullName)
    sheet['B21'].value = ('These tabs show positions with more than {} unique'
                          ' Your Titles'.format(int(tresholds.iloc[6,1])))
    sheet['D32'].value = SubmissionDeadline
    sheet['D33'].value = InternalDeadline
    sheet['D34'].value = InternalDeadline
    sheet['D35'].value = QAD

    # Updating Summary & QAD sheet with treshold value
    sheet = template['Summary']
    sheet['P1'] = 'DVF more than {} days after QAD'.format(int(tresholds.iloc[7,1]))
    sheet = template['DVF after QAD & MDA']
    sheet['K4'] = 'More than {} days after QAD'.format(int(tresholds.iloc[7,1]))

    # copying treshold values, with an addition of formatting
    sheet = template['Tresholds']
    sheet['C2'] = '>{}%'.format(int(tresholds.iloc[0,1]*100))
    sheet['C3'] = '>{}'.format(int(tresholds.iloc[1,1]))
    sheet['C4'] = '>{}%'.format(int(tresholds.iloc[2,1]*100))
    sheet['C5'] = '>{}'.format(int(tresholds.iloc[3,1]))
    sheet['C6'] = '<{}%'.format(int(tresholds.iloc[4,1]*100))
    sheet['C7'] = '>{}'.format(int(tresholds.iloc[5,1]))
    sheet['C8'] = '>{}'.format(int(tresholds.iloc[6,1]))
    sheet['C9'] = '>{}'.format(int(tresholds.iloc[7,1]))

    # saving the file in the working folder and with final name
    print('saving template...')
    template.save('{}\\{}'.format(Path, ReportName))




#   CREATING FINAL REPORT
#   functions for getting final data for our template
#   and saving them in final Excel
def create_final_report():

    # creating writer tool which will copy & paste data
    # while preserving original file structure
    print('preparing to copy data...')
    Report = '{}\\{}'.format(Path, ReportName)
    Writer = pd.ExcelWriter(Report, engine='openpyxl')
    Book = pyxl.load_workbook(Report)
    Writer.book = Book
    Writer.sheets = dict((ws.title, ws) for ws in Book.worksheets)

    # creating summary
    print('copying Summary data...')
    Summary.iloc[:, 0:7].to_excel(Writer,
                                  sheet_name='Summary',
                                  header=None, index=False,
                                  startrow=1)
    Summary.iloc[:, 7:].to_excel(Writer,
                                 sheet_name='Summary',
                                 header=None, index=False,
                                 startrow=1, startcol=8)


    # checking each tab for YN confirmation for tool to copy its data
    if ExcEntities_YN == 'Y':
        print('copying Excluded Entities data...')
        ExcEntitiesData.iloc[:, 0:7].to_excel(Writer,
                                              sheet_name='Excluded Entities',
                                              header=None, index=False,
                                              startrow=4)
        ExcEntitiesData.iloc[:, 7:].to_excel(Writer,
                                             sheet_name='Excluded Entities',
                                             header=None, index=False,
                                             startrow=4, startcol=8)
    else:
        print('skipping Excluded entities')


    if ExcIncumbents_YN == 'Y':
        print('copying Excluded Incumbents data...')
        ExcIncumbentsData.iloc[:, 0:8].to_excel(Writer,
                                                sheet_name='Excluded Incumbents',
                                                header=None, index=False,
                                                startrow=4)
        ExcIncumbentsData.iloc[:, 8:].to_excel(Writer,
                                               sheet_name='Excluded Incumbents',
                                               header=None, index=False,
                                               startrow=4, startcol=9)
    else:
        print('skipping Excluded Incumbents')


    if AGCRematches_YN == 'Y':
        print('copying AGC Rematches data...')
        AGCRematchesData.iloc[:, 0:8].to_excel(Writer,
                                               sheet_name='AGC Rematches',
                                               header=None, index=False,
                                               startrow=4)
        AGCRematchesData.iloc[:, 8:].to_excel(Writer,
                                              sheet_name='AGC Rematches',
                                              header=None, index=False,
                                              startrow=4, startcol=9)
    else:
        print('skipping AGC Rematches')


    if MultipleYTPos_YN == 'Y':
        print('copying Multiple YT Positions data...')
        MultipleYTPositionsData.iloc[:, 0:8].to_excel(Writer,
                                                      sheet_name='Multiple YT Positions',
                                                      header=None, index=False,
                                                      startrow=4)
        MultipleYTPositionsData.iloc[:, 8:].to_excel(Writer,
                                                     sheet_name='Multiple YT Positions',
                                                     header=None, index=False,
                                                     startrow=4, startcol=9)
    else:
        print('skipping Multiple YT Positions')


    if MultipleYTAll_YN == 'Y':
        print('copying Multiple YT All Data...')
        MultipleYTPositionsAll.iloc[:, 0:8].to_excel(Writer,
                                                     sheet_name='Multiple YT All Data',
                                                     header=None, index=False,
                                                     startrow=4)
        MultipleYTPositionsAll.iloc[:, 8:].to_excel(Writer,
                                                    sheet_name='Multiple YT All Data',
                                                    header=None, index=False,
                                                    startrow=4, startcol=9)
    else:
        print('skipping Multiple YT All Data')


    if Deadlines_YN == 'Y':
        print('copying After Deadlines data...')
        AfterDeadlinesData.iloc[:, 0:7].to_excel(Writer,
                                                 sheet_name='After Deadline',
                                                 header=None, index=False,
                                                 startrow=4)
        AfterDeadlinesData.iloc[:, 7:].to_excel(Writer,
                                                sheet_name='After Deadline',
                                                header=None, index=False,
                                                startrow=4, startcol=8)
    else:
        print('skipping After Deadlines data')


    if QAD_YN == 'Y':
        print('copying QAD & MDA data...')
        MDAData.iloc[:, 0:7].to_excel(Writer,
                                      sheet_name='DVF after QAD & MDA',
                                      header=None, index=False,
                                      startrow=4)
        MDAData.iloc[:, 7:].to_excel(Writer,
                                     sheet_name='DVF after QAD & MDA',
                                     header=None, index=False,
                                     startrow=4, startcol=8)
    else:
        print('skipping QAD & MDA data')

    print('saving...')
    Writer.save()
    print('\nReport is ready at {}.'.format(Path))




# nice header to display in conda prompt
def intro_header():

    print('     ____  ______   ______            __              ')
    print('    / __ \/ ____/  /_  __/___  ____  / /              ')
    print('   / /_/ / /        / / / __ \/ __ \/ /               ')
    print('  / ____/ /___     / / / /_/ / /_/ / /                ')
    print(' /_/    \____/    /_/  \____/\____/_/   version 0.99\n')




# defining main function connecting all minor steps:
def __main__():

    intro_header()
    create_tresholds()
    import_xlsx_files(Path)
    make_names_standard()
    create_excluded_entities()
    create_exclude_reasons()
    create_excluded_incumbents()
    create_YTPos()
    create_YTAll()
    create_after_deadline()
    create_mda()
    create_contact_details()
    create_summary()
    create_latesub()
    create_report_file()
    create_final_report()

__main__()
