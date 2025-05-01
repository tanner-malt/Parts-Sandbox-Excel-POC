'''
This file is the Manager for the Parts Sandbox application.
It is responsible for managing the overall flow of the application, including GUI, handling user interactions, and acting as the "Owner" of the excel database.
'''
import openpyxl
import pandas as pd


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

    def update_alias(self, file_path):
        '''
        Updates the alias in the database.

        This function processes the "Master Part List" sheet in the provided Excel file and updates the "aliases" sheet 
        in the `parts_sandbox.xlsx` database. It ensures that aliases from the "Master Part List" are added to the 
        "aliases" sheet only if they do not already exist.

        Parameters:
            file_path (str): The path to the Excel file containing the "Master Part List" sheet.

        Functionality:
            - Reads the "Master Part List" sheet from the provided Excel file.
            - Checks if each alias in the "Master Part List" already exists in the "aliases" sheet of the database.
            - If an alias does not exist, it adds the alias and its corresponding value to the "aliases" sheet.
            - The "aliases" sheet acts as a simple alias-to-value mapping.

        Assumptions:
            - The "Master Part List" sheet contains aliases in column A and their corresponding values in column B.
            - The "aliases" sheet in the `parts_sandbox.xlsx` file has the same structure: aliases in column A and values in column B.

        Raises:
            FileNotFoundError: If the provided file or the database file does not exist.
            KeyError: If the "Master Part List" sheet is missing in the provided file.
            Exception: For any other unexpected errors.

        Example:
            Given a "Master Part List" sheet with the following data:
                | Alias       | Value       |
                |-------------|-------------|
                | Part123     | Widget A    |
                | Part456     | Widget B    |

            And an "aliases" sheet with the following data:
                | Alias       | Value       |
                |-------------|-------------|
                | Part123     | Widget A    |

            After running this function, the "aliases" sheet will be updated to:
                | Alias       | Value       |
                |-------------|-------------|
                | Part123     | Widget A    |
                | Part456     | Widget B    |
        '''
        try:
            print(f'Loading workbook at {file_path}')
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook['Master Part List']

            #Make a dataframe of the sheet
            df = pd.DataFrame(sheet.values)
            print(df.head())
        except FileNotFoundError:
            raise FileNotFoundError(f"File {file_path} not found.")