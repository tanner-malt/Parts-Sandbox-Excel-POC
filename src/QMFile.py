'''
This file contains the Quote Master File class and its associated methods.
This class is responsible for managing the Quote Master File, including loading, saving, and manipulating the data within it.
'''
import openpyxl

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
            print("Trying")
            self.file_path = file_path
            self.workbook = openpyxl.load_workbook(file_path)
            self.data = self.load_data()
        except FileNotFoundError:
            raise FileNotFoundError(f"File {file_path} not found.")
            del self
