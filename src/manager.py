'''
This file is the Manager for the Parts Sandbox application.
It is responsible for managing the overall flow of the application, including GUI, handling user interactions, and acting as the "Owner" of the excel database.
'''
import openpyxl

class ApplicationManager:
    '''
    This class is the manager for the Parts Sandbox application.
    It is responsible for managing the overall flow of the application, including GUI, handling user interactions, and acting as the "Owner" of the excel database.
    '''
    def __init__(self):
        '''
        Initializes the application manager.
        '''
        self.start()

    def start(self):
        '''
        Starts the application manager.
        '''
        self.check_for_db()

    def check_for_db(self):
        '''
        Checks for the database file.
        If it does not exist, it creates a new one.
        '''
        try:
            print("Trying to load the workbook", flush=True)
            self.workbook = openpyxl.load_workbook(r'excels/parts_sandbox.xlsx')
        except FileNotFoundError:
            print("Database file not found. Creating a new one.", flush=True)
            workbook = openpyxl.Workbook()
            workbook.save(r'excels/parts_sandbox.xlsx')