import numpy as np
import pandas as pd
import csv
import re
from pandas import ExcelWriter
from pandas import ExcelFile
import xlsxwriter
import sys
import time
import keyboard
import threading
import os
import shutil
from tkinter import Tk
from tkinter.filedialog import askdirectory
import PySimpleGUI as sg
import pickle
from typing import Dict

sg.ChangeLookAndFeel('Black')

''' Create windows base GUI object '''
class Gui:
    def __init__(self):
        self.layout: list = \
            [
                [sg.Text('Source Folder', size=(15, 1), auto_size_text=False, justification='left',
                         tooltip='Target folder to search for the specified file...'),
                 sg.InputText('', size=(75, 1), key='SOURCE_FOLDER_PATH', text_color='Black', disabled=True,
                              tooltip='Target folder to search for the specified file...'),
                 sg.FolderBrowse('Browse', size=(6, 1), key='_BROWSE_SOURCE_FOLDER_')],
                [sg.Text('Destination Folder', size=(15, 1), auto_size_text=False, justification='left',
                         tooltip='Target storage folder to store the copies of matched file...'),
                 sg.InputText('', size=(75, 1), key='DESTINATION_FOLDER_PATH', text_color='Black', disabled=True,
                              tooltip='Target storage folder to store the copies of matched file...'),
                 sg.FolderBrowse('Browse', size=(6, 1), key='_BROWSE_DESTINATION_FOLDER_')],
                [sg.Text('_' * 115, size=(90, 1))],
                [sg.Radio('Segregate Auto Test 3 Phase Waveform', size=(35, 1), group_id='choice_type',
                          key="_FIND_WAVEFORM_",
                          enable_events=True,
                          tooltip='Segregate Waveform .png files by reference, pass, fail, test case '
                                  'base on output cleaned .xlsx file of the automatic tester test result',
                          default=True),
                 sg.Radio('Search by File List', size=(20, 1), group_id='choice_type',
                          key="_FIND_FILE_LIST_",
                          enable_events=True,
                          tooltip='Search multiple file name using excel list'),
                 sg.Radio('Search by File Name', size=(20, 1), group_id='choice_type',
                          key="_FIND_FILE_NAME_",
                          enable_events=True,
                          tooltip='Search single file name')],
                [sg.Text('Cleaned .xlsx File', size=(15, 1), auto_size_text=False, justification='left',
                         tooltip='Select cleaned output file generated from 3PTapDataCleaningVxx'),
                 sg.InputText('', size=(75, 1), key='WAVEFORM_LIST', text_color='Black', disabled=True,
                              tooltip='Select cleaned output file generated from 3PTapDataCleaningVxx'),
                 sg.FileBrowse('Browse', size=(6, 1), key='_BROWSE_LIST_', file_types=(("Excel Files", "*.xlsx"),))],
                [sg.Text('_' * 115, size=(90, 1))],
                [sg.Radio('Contain', size=(18, 1), group_id='subchoice', key="_CONTAIN_", disabled=True,
                          enable_events=True,
                          tooltip='Search files with filename contain text input in <Searched Filename> field'),
                 sg.Radio('Begin With', size=(18, 1), group_id='subchoice', key="_BEGIN_WITH_", disabled=True,
                          enable_events=True, default=True,
                          tooltip='Search files with filename begins with text input in <Searched Filename> field'),
                 sg.Radio('End With', size=(18, 1), group_id='subchoice', key="_END_WITH_", disabled=True,
                          enable_events=True,
                          tooltip='Search files with filename end with text input in <Searched Filename> field')],
                [sg.Text('Searched Filelist', size=(15, 1), auto_size_text=False, justification='left',
                         tooltip='Select .xlsx file with list of filename located in column A without title head'),
                 sg.InputText('', size=(75, 1), key='FILELIST', text_color='Black', disabled=True,
                              tooltip='Select .xlsx file with list of filename located in column A without title head'),
                 sg.FileBrowse('Browse', size=(6, 1), key='_BROWSE_FILELIST_', file_types=(("Excel Files", "*.xlsx"),),
                               disabled=True)],
                [sg.Text('Searched Filename', size=(15, 1), auto_size_text=False, justification='left',
                         tooltip='Key in partial or full specified file name for searching...'),
                 sg.InputText('', size=(60, 1), key='FILENAME', text_color='White', disabled=True,
                              tooltip='Key in partial or full specified file name for searching...'),
                 sg.Ok(size=(11, 1), key='_OK BUTTON_', button_color=('white', 'blue')),
                 sg.Cancel(size=(6, 1), key='_CANCEL BUTTON_')],
                [sg.Output(size=(100, 20))]
            ]
        self.window: object = sg.Window(
            'Find Files in Source Folder and Copy To Destination Folder',
            self.layout,
            auto_size_text=True,
            element_justification='left')


''' Create file search engine object '''
class SearchEngine:
    def __init__(self):
        self.file_index = []  # directory listing returned by os.walk()
        self.results = []  # search results returned from search method
        self.matches = 0  # count of records matched
        self.records = 0  # count of records searched

    ''' Create a new file index of the root; then save to self.file_index and to pickle file '''
    def create_new_index(self, values: Dict[str, str]) -> None:
        source_root_path = values['SOURCE_FOLDER_PATH']
        destination_root_path = values['DESTINATION_FOLDER_PATH']
        self.file_index: list = [(root, files) for root, dirs, files in os.walk(source_root_path) if files]

        ''' Save index to file '''
        file_index_path = os.path.join(destination_root_path, 'file_index.pkl')
        with open(file_index_path, 'wb') as f:
            pickle.dump(self.file_index, f)

    ''' Load existing file index into program '''
    def load_existing_index(self, values: Dict[str, str]) -> None:
        source_root_path = values['SOURCE_FOLDER_PATH']
        destination_root_path = values['DESTINATION_FOLDER_PATH']
        try:
            file_index_path = os.path.join(destination_root_path, 'file_index.pkl')
            with open(file_index_path, 'rb') as f:
                self.file_index = pickle.load(f)
        except:
            self.file_index = []

    ''' Delete existing file index from folder '''
    def delete_existing_index(self, values: Dict[str, str]) -> None:
        source_root_path = values['SOURCE_FOLDER_PATH']
        destination_root_path = values['DESTINATION_FOLDER_PATH']
        try:
            file_index_path = os.path.join(destination_root_path, 'file_index.pkl')
            os.remove(file_index_path)
        except:
            pass

    ''' Search for the term base on the type in the index; the types of search include: 
        contain, begin with, end with; save the result to file '''
    def search(self, values: Dict[str, str]) -> None:
        source_root_path = values['SOURCE_FOLDER_PATH']
        destination_root_path = values['DESTINATION_FOLDER_PATH']
        self.results.clear()
        self.matches = 0
        self.records = 0
        term = str(values['FILENAME'])

        # search for matches and count results
        for path, files in self.file_index:
            for file in files:
                self.records += 1
                if (values['_CONTAIN_'] and term.lower() in file.lower()
                        or values['_BEGIN_WITH_'] and file.lower().startswith(term.lower())
                        or values['_END_WITH_'] and file.lower().endswith(term.lower())):
                    result = path.replace('\\', '/') + '/' + file
                    shutil.copy(result, destination_root_path)
                    self.results.append(result)
                    self.matches += 1
                else:
                    continue

        # save results to file
        result_path = os.path.join(destination_root_path, 'search_results.txt')
        with open(result_path, 'w') as f:
            for row in self.results:
                f.write(row + '\n')


def OneSetStrMatch(Arr1, Arr1_Value, Reg1):
    return Arr1.str.contains(pat=Arr1_Value, regex=Reg1)


def SearchAndCopy(s, values):
    s.create_new_index(values)
    print()
    print(">> Index file created...")
    print(">> Searching file using just created index file...")
    s.search(values)
    print()
    for result in s.results:
        print(result)
    print()
    print(">> Searched {:,d} records and found {:,d} matches".format(s.records, s.matches))
    print(">> Results saved in working directory as search_results.txt.")
    return s.matches

''' main loop of the program '''
def main():
    g = Gui()
    s = SearchEngine()
    #s.load_existing_index()  # load if exists, otherwise return empty list

    while True:
        event, values = g.window.read()

        if event == '_FIND_WAVEFORM_':
            g.window['_BROWSE_LIST_'].Update(disabled=False)
            g.window['WAVEFORM_LIST'].Update(text_color='Black')
            g.window['_BROWSE_FILELIST_'].Update(disabled=True)
            g.window['FILELIST'].Update(text_color='LightGrey')
            g.window['FILENAME'].Update(text_color='LightGrey', disabled=True)
            g.window['_CONTAIN_'].Update(disabled=True)
            g.window['_BEGIN_WITH_'].Update(disabled=True)
            g.window['_END_WITH_'].Update(disabled=True)
            #print(values['_CONTAIN_'], values['_BEGIN_WITH_'], values['_END_WITH_'])
            #print('_FIND_WAVEFORM_')

        if event == '_FIND_FILE_LIST_':
            g.window['_BROWSE_LIST_'].Update(disabled=True)
            g.window['WAVEFORM_LIST'].Update(text_color='LightGrey')
            g.window['_BROWSE_FILELIST_'].Update(disabled=False)
            g.window['FILELIST'].Update(text_color='Black')
            g.window['FILENAME'].Update(text_color='LightGrey', disabled=True)
            g.window['_CONTAIN_'].Update(disabled=False)
            g.window['_BEGIN_WITH_'].Update(disabled=False)
            g.window['_END_WITH_'].Update(disabled=False)
            #print(values['_CONTAIN_'], values['_BEGIN_WITH_'], values['_END_WITH_'])
            #print('_FIND_FILE_LIST_')

        if event == '_FIND_FILE_NAME_':
            g.window['_BROWSE_LIST_'].Update(disabled=True)
            g.window['WAVEFORM_LIST'].Update(text_color='LightGrey')
            g.window['_BROWSE_FILELIST_'].Update(disabled=True)
            g.window['FILELIST'].Update(text_color='LightGrey')
            g.window['FILENAME'].Update(text_color='White', disabled=False)
            g.window['_CONTAIN_'].Update(disabled=False)
            g.window['_BEGIN_WITH_'].Update(disabled=False)
            g.window['_END_WITH_'].Update(disabled=False)
            #print(values['_CONTAIN_'], values['_BEGIN_WITH_'], values['_END_WITH_'])
            #print('_FIND_FILE_NAME_')

        #if event in ('_CONTAIN_', '_BEGIN_WITH_', '_END_WITH_'):
            #print(values['_CONTAIN_'], values['_BEGIN_WITH_'], values['_END_WITH_'])
            #g.window['_BROWSE_LIST_'].Update(disabled=True)
            #g.window['WAVEFORM_LIST'].Update(text_color='DarkGrey')
            #g.window['FILENAME'].Update(text_color='White', disabled=False)

        if event in (None, '_CANCEL BUTTON_'):
            print(">> Program is aborted")
            #time.sleep(3)
            break

        if event == '_OK BUTTON_':
            if values['_FIND_WAVEFORM_'] == True:
                if (values['SOURCE_FOLDER_PATH'] == '') | (values['DESTINATION_FOLDER_PATH'] == '') | (values['WAVEFORM_LIST'] == ''):
                    print(">> Source Folder, Destination Folder or .xlsx File List fields cannot be left empty.")
                else:
                    ''' Extract column from excel '''
                    colName = ['Deleted', 'ID', 'Test_Case_Description',
                               'Reference', 'Product_Serial_Number',
                               'Trigger_Result', 'Recovery_Result']

                    print(">> Reading excel file...")
                    print(">> Time taken to complete reading will be longer with increasing of file size...")
                    print(">> Please wait until <End of reading excel file...> message appears...")
                    dfFileList = pd.read_excel(values['WAVEFORM_LIST'],
                                               sheet_name='placeholderSelectedQuery',
                                               names=colName,
                                               usecols='A,C,K,X,Z,IQ,IR')

                    ''' Combined Reference & Product_Serial_Number '''
                    dfFileList['Ref_Serial'] = dfFileList['Reference'] + '_' + dfFileList['Product_Serial_Number']

                    ''' Remove specific character in the test case description'''
                    dfFileList['Test_Case_Description'] = dfFileList['Test_Case_Description'].str.replace('[>%/]', '')

                    ''' Convert ID column from Int to String '''
                    dfFileList['ID'] = dfFileList['ID'].apply(str)

                    ''' Fill NaN value with empty '''
                    dfFileList['Trigger_Result'].fillna('', inplace=True)
                    dfFileList['Recovery_Result'].fillna('', inplace=True)

                    ''' Shift hysteresis result down 1 cell '''
                    dfHYS = dfFileList.loc[:, ['Recovery_Result']]
                    dfHYS = dfHYS.shift(1, axis=0)  # Shift vertically down by 1 index (1st row becomes NaN)
                    dfFileList['HYS'] = dfHYS

                    dfFileList['Temp'] = np.where(OneSetStrMatch(dfFileList['Test_Case_Description'],
                                                                           '^.*(HYS).*$', True),
                                                   dfFileList['HYS'],
                                                   dfFileList['Trigger_Result'])

                    dfFileList['Trigger_Result'] = dfFileList['Temp']

                    #CleanedCSV = os.path.join(values['DESTINATION_FOLDER_PATH'], "Cleaned_1" + ".csv")
                    # Exporting cleaned data in CSV format for each worksheet data
                    #dfFileList.to_csv(CleanedCSV, index=False, index_label=False)

                    ''' Remove rows '''
                    #dfFileList.drop(dfFileList[dfFileList['Test_Case_Description'].str.contains(r'HYS')].index, inplace=True)
                    #dfFileList.drop(dfFileList.loc[dfFileList['Deleted'] == 1].index, inplace=True)

                    ''' Remove cols '''
                    dfFileList = dfFileList.drop(['HYS', 'Temp'], axis=1)

                    #CleanedCSV = os.path.join(values['DESTINATION_FOLDER_PATH'], "Cleaned_2" + ".csv")
                    # Exporting cleaned data in CSV format for each worksheet data
                    #dfFileList.to_csv(CleanedCSV, index=False, index_label=False)

                    ''' Get unique value of each column '''
                    dfRefSerial = pd.unique(dfFileList['Ref_Serial']).tolist()
                    dfTriggerResult = pd.unique(dfFileList['Trigger_Result']).tolist()
                    dfRecoveryResult = pd.unique(dfFileList['Recovery_Result']).tolist()
                    dfTestCase = pd.unique(dfFileList['Test_Case_Description']).tolist()

                    print(">> End of reading excel file...")
                    #print(dfFileList.info(verbose=True))
                    #print(dfFileList.dtypes)
                    #print(dfFileList.head(5))
                    #print(dfFileList.tail(5))

                    ''' Create directories or folders '''
                    Total_Iteration = len(dfRefSerial) * len(dfTriggerResult) * len(dfTestCase)
                    z = 0
                    for i in dfRefSerial:
                        for j in dfTriggerResult:
                            for k in dfTestCase:
                                z = z + 1
                                ContinueSeach = sg.one_line_progress_meter('Creating Folder',
                                                                              z,
                                                                              Total_Iteration,
                                                                              'key',
                                                                              'Creating folders ...',
                                                                              orientation='horizontal')

                                directory_path = os.path.join(values['DESTINATION_FOLDER_PATH'], i, j, k)
                                try:
                                    os.makedirs(directory_path)
                                except OSError:
                                    if not os.path.isdir(directory_path):
                                        raise

                                if ContinueSeach == False:
                                    break

                    ''' Update window keys & call searching & copy function '''
                    results =[]
                    Total_Iteration = len(dfFileList['ID'])
                    z = 0
                    for r in zip(dfFileList['ID'],
                                 dfFileList['Ref_Serial'],
                                 dfFileList['Trigger_Result'],
                                 dfFileList['Test_Case_Description']):

                        z = z + 1
                        ContinueSeach = sg.one_line_progress_meter('Sorting Waveform Image',
                                                                      z,
                                                                      Total_Iteration,
                                                                      'key',
                                                                      'Sorting Waveform .png file into folders ...',
                                                                      orientation='horizontal')

                        destination_path = os.path.join(values['DESTINATION_FOLDER_PATH'], r[1], r[2], r[3])
                        Temp_Dict = {'SOURCE_FOLDER_PATH': values['SOURCE_FOLDER_PATH'],
                                     'DESTINATION_FOLDER_PATH': destination_path,
                                     'FILENAME': r[0] + '_',
                                     '_CONTAIN_': False,
                                     '_BEGIN_WITH_': True,
                                     '_END_WITH_': False}

                        #print(Temp_Dict['SOURCE_FOLDER_PATH'], " | ",
                        #      Temp_Dict['DESTINATION_FOLDER_PATH'], " | ",
                        #      Temp_Dict['FILENAME'], " | ",
                        #      Temp_Dict['_CONTAIN_'], " | ",
                        #      Temp_Dict['_BEGIN_WITH_'], " | ",
                        #      Temp_Dict['_END_WITH_'])

                        No_of_Match = SearchAndCopy(s, Temp_Dict)
                        result = [Temp_Dict['FILENAME'], r[1], r[2], r[3], No_of_Match]
                        results.append(result)

                        if ContinueSeach == False:
                            break

                    # save results to file as .csv
                    result_path = os.path.join(values['DESTINATION_FOLDER_PATH'], 'search_results.csv')
                    with open(result_path, 'w') as ofile:
                        writer = csv.writer(ofile)
                        writer.writerows(results)

            elif values['_FIND_FILE_LIST_'] == True:
                if (values['SOURCE_FOLDER_PATH'] == '') | (values['DESTINATION_FOLDER_PATH'] == '') | (values['FILELIST'] == ''):
                    print(">> Source Folder, Destination Folder or Search Filelist fields cannot be left empty.")
                else:
                    ''' Extract column from excel '''
                    colName = ['Filelist']

                    ''' Extract column from excel '''
                    print(">> Reading excel file...")
                    print(">> Time taken to complete reading will be longer with increasing of file size...")
                    print(">> Please wait until <End of reading excel file...> message appears...")
                    dfFileList = pd.read_excel(values['FILELIST'],
                                               sheet_name=0,
                                               header=None,
                                               names=colName,
                                               usecols='A',
                                               dtype={'Filelist': str})
                    print(">> End of reading excel file...")

                    ''' Delete empty row in filelist '''
                    dfFileList = dfFileList.dropna()

                    ''' Update window keys & call searching & copy function '''
                    results = []
                    Total_Iteration = len(dfFileList['Filelist'])
                    z = 0
                    for r in zip(dfFileList['Filelist']):
                        z = z + 1
                        ContinueSeach = sg.one_line_progress_meter('Searching file',
                                                                   z,
                                                                   Total_Iteration,
                                                                   'key',
                                                                   'Searching file base on file lists ...',
                                                                   orientation='horizontal')

                        destination_path = os.path.join(values['DESTINATION_FOLDER_PATH'])
                        Temp_Dict = {'SOURCE_FOLDER_PATH': values['SOURCE_FOLDER_PATH'],
                                     'DESTINATION_FOLDER_PATH': destination_path,
                                     'FILENAME': r[0],
                                     '_CONTAIN_': values['_CONTAIN_'],
                                     '_BEGIN_WITH_': values['_BEGIN_WITH_'],
                                     '_END_WITH_': values['_END_WITH_']}

                        # print(Temp_Dict['SOURCE_FOLDER_PATH'], " | ",
                        #      Temp_Dict['DESTINATION_FOLDER_PATH'], " | ",
                        #      Temp_Dict['FILENAME'], " | ",
                        #      Temp_Dict['_CONTAIN_'], " | ",
                        #      Temp_Dict['_BEGIN_WITH_'], " | ",
                        #      Temp_Dict['_END_WITH_'])

                        No_of_Match = SearchAndCopy(s, Temp_Dict)
                        result = [Temp_Dict['FILENAME'], No_of_Match]
                        results.append(result)

                        if ContinueSeach == False:
                            break

                    # save results to file as .csv
                    result_path = os.path.join(values['DESTINATION_FOLDER_PATH'], 'search_results.csv')
                    with open(result_path, 'w') as ofile:
                        writer = csv.writer(ofile)
                        writer.writerows(results)

            else:
                if (values['SOURCE_FOLDER_PATH'] == '') | (values['DESTINATION_FOLDER_PATH'] == '') | (values['FILENAME'] == ''):
                    print(">> Source Folder, Destination Folder or Search Filename fields cannot be left empty.")
                else:
                    No_of_Match = SearchAndCopy(s, values)

    g.window.close()

if __name__ == '__main__':
    start = time.time()  # Record start time
    print("<< Starting program... by", os.getlogin(), "under " + sys.platform + " operating system >>")
    main()


