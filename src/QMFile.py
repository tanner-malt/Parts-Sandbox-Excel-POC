'''
This file contains the Quote Master File class and its associated methods.
This class is responsible for managing the Quote Master File, including loading, saving, and manipulating the data within it.
'''
import openpyxl
import pandas as pd

class QMFile:
    '''
    This class represents one of many Quote Master Files.
    It is responsible for being the python interface to the Quote Master File.
    '''
    def __init__(self, file_path):
        '''
        Initializes the QMFile object with the given file path.
        Loads the workbook and initializes the data structure.
        '''
        try:
            self.file_path = file_path
            self.workbook = openpyxl.load_workbook(file_path)
            self.data = self.load_data()
        except FileNotFoundError:
            raise FileNotFoundError(f"File {file_path} not found.")
            del self

    def load_master_sheet(self):
        '''
        Loads the master sheet from the workbook.
        Returns the master sheet object.
        '''
        try:
            return self.workbook['Master Part List']
        except KeyError:
            raise KeyError("Master Part List sheet not found in the workbook.")

    def load_data(self):
        '''
        Loads data from the workbook into a structured format.
        Returns a dictionary containing the parsed data.
        '''
        try:
            master_sheet = self.load_master_sheet()
            df = pd.DataFrame(master_sheet.values)
            # Assuming first row contains headers
            df.columns = df.iloc[0]
            df = df.iloc[1:]  # Remove header row from data
            return {
                'parts': df.to_dict('records')
            }
        except Exception as e:
            raise Exception(f"Error loading data: {str(e)}")
