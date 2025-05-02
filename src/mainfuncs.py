'''
File to define the main functions for the project.
'''
import os
import openpyxl
import GUI

def list_quote_master_files():
    """
    Lists all Quote Master Files in the specified directory.
    Eventually will need to be replaced with a more robust file management system and path viewer, fine for demo.
    This function retrieves all files with the .xlsx extension in the specified directory.
    """
    # Define the directory containing the Quote Master Files
    quote_master_files_directory = "excels"
    
    # List all files in the directory
    try:
        files = os.listdir(quote_master_files_directory)
        quote_master_files = [file for file in files if file.endswith('.xlsx')]

        if len(quote_master_files) == 0:
            print("No Quote Master Files found.")
            return
        print("Quote Master Files:")

        for file in quote_master_files:
            print(file)
            
        return quote_master_files
    except FileNotFoundError:
        print(f"Directory {quote_master_files_directory} not found.")
        return []
    
def prepare_master(master_sheet):
    """
    Prepares the master sheet for processing.
    """


def analysis_loop():
    '''
    This is the analysis loop called to facilitate planning, things such as looking up MOQ, total spend, etc can be found here.
    '''
    workbook = openpyxl.load_workbook("excels\parts_sandbox.xlsx")
    GUI.make_analysis_window()
    
    print(workbook)