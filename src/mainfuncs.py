'''
File to define the main functions for the project.
'''
import os
import openpyxl
import GUI
from QMFile import QMFile
from manager import ApplicationManager
import pandas as pd

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
    
def prepare_master(file_path=None):
    """
    Prepares the master sheet for processing by updating the database.
    If no file_path is provided, processes all Quote Master files in the excels directory.
    
    Args:
        file_path (str, optional): Path to a specific Quote Master file to process
    """
    try:
        # Create an instance of ApplicationManager
        app_manager = ApplicationManager()
        sandbox_path = os.path.join("excels", "parts_sandbox.xlsx")
        
        # If no specific file provided, process all files
        if file_path is None:
            excel_folder = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'excels'))
            files = [
                os.path.join(excel_folder, f) 
                for f in os.listdir(excel_folder) 
                if f.endswith('.xlsx') and f != "parts_sandbox.xlsx"
            ]
        else:
            files = [file_path]
            
        # Load existing aliases from parts_sandbox.xlsx
        try:
            existing_aliases = pd.read_excel(sandbox_path, sheet_name='aliases')
        except (FileNotFoundError, KeyError):
            # Create new aliases sheet if it doesn't exist
            existing_aliases = pd.DataFrame(columns=['alias', 'value'])
            with pd.ExcelWriter(sandbox_path, engine='openpyxl', mode='a' if os.path.exists(sandbox_path) else 'w') as writer:
                existing_aliases.to_excel(writer, sheet_name='aliases', index=False)
        
        # Process each file
        all_aliases = []
        for file in files:
            try:
                qm_file = QMFile(file)
                master_sheet = qm_file.load_master_sheet()
                df = pd.DataFrame(master_sheet.values)
                # Skip header row and get alias/value columns
                df.columns = df.iloc[0]
                df = df.iloc[1:]
                if 'alias' in df.columns and 'value' in df.columns:
                    aliases_from_file = df[['alias', 'value']].copy()
                    all_aliases.append(aliases_from_file)
            except Exception as e:
                print(f"Error processing {file}: {str(e)}")
                continue
        
        if all_aliases:
            # Combine all new aliases
            new_aliases = pd.concat(all_aliases, ignore_index=True)
            # Remove duplicates, keeping first occurrence
            new_aliases = new_aliases.drop_duplicates(subset=['alias'], keep='first')
            
            # Merge with existing aliases, keeping existing ones in case of conflict
            combined_aliases = pd.concat([existing_aliases, new_aliases], ignore_index=True)
            combined_aliases = combined_aliases.drop_duplicates(subset=['alias'], keep='first')
            
            # Save back to parts_sandbox.xlsx
            with pd.ExcelWriter(sandbox_path, engine='openpyxl', mode='a') as writer:
                combined_aliases.to_excel(writer, sheet_name='aliases', index=False)
            
            return True
            
    except Exception as e:
        raise Exception(f"Error preparing master file: {str(e)}")

def analysis_loop():
    '''
    This is the analysis loop called to facilitate planning, things such as looking up MOQ, total spend, etc can be found here.
    '''
    sandbox_path = os.path.join("excels", "parts_sandbox.xlsx")
    workbook = openpyxl.load_workbook(sandbox_path)
    GUI.make_analysis_window()
    
    print(workbook)