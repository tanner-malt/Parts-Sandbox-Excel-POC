'''
This file is the Manager for the Parts Sandbox application.
It is responsible for managing the overall flow of the application, including database operations and request handling.
'''
import sqlite3
import pandas as pd
import os
import logging
from QMFile import QMFile

logger = logging.getLogger(__name__)

class ApplicationManager:
    '''
    This class is the manager for the Parts Sandbox application.
    It handles database operations and serves as the backend for the application.
    '''
    def __init__(self):
        '''
        Initializes the application manager and database.
        '''
        self.db_path = 'database/parts_sandbox.db'
        self.excel_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'excels'))
        self._conn = None
        self.start()

    @contextmanager
    def get_connection(self):
        """Get a database connection within a context"""
        if ':memory:' in self.db_path:
            # For in-memory database, maintain the same connection
            if self._conn is None:
                self._conn = sqlite3.connect(self.db_path)
                self._init_db(self._conn)
            yield self._conn
        else:
            # For file-based database, create new connection each time
            conn = sqlite3.connect(self.db_path)
            try:
                yield conn
                conn.commit()
            finally:
                conn.close()

    def _init_db(self, conn):
        """Initialize database tables"""
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS aliases (
                alias TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS quote_masters (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT NOT NULL,
                last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        conn.commit()

    def start(self):
        '''
        Starts the application manager and ensures database exists.
        '''
        logger.info("Starting ApplicationManager")
        self.check_for_db()
        
    def check_for_db(self):
        '''
        Checks for the database file.
        If it does not exist, it creates a new one with necessary tables.
        '''
        try:
            logger.info("Checking for parts_sandbox.xlsx")
            self.workbook = openpyxl.load_workbook(r'excels/parts_sandbox.xlsx')
        except FileNotFoundError:
            logger.info("Database file not found. Creating a new one.")
            workbook = openpyxl.Workbook()
            workbook.save(r'excels/parts_sandbox.xlsx')

    def update_alias(self, file_path):
        '''
        Updates the alias in the database from an Excel file.
        '''
        try:
            logger.info(f"Updating aliases from file: {file_path}")
            df = pd.read_excel(file_path, sheet_name='Master Part List')
            
            if 'alias' in df.columns and 'value' in df.columns:
                with self.get_connection() as conn:
                    # Create temporary table for new aliases
                    cursor = conn.cursor()
                    cursor.execute('''
                        CREATE TABLE IF NOT EXISTS temp_aliases (
                            alias TEXT PRIMARY KEY,
                            value TEXT NOT NULL
                        )
                    ''')
                    
                    # Convert DataFrame to SQL
                    alias_data = df[['alias', 'value']].copy()
                    alias_data.to_sql('temp_aliases', conn, if_exists='replace', index=False)
                    
                    # Merge new aliases with existing ones
                    cursor.execute('''
                        INSERT OR REPLACE INTO aliases (alias, value)
                        SELECT alias, value FROM temp_aliases
                    ''')
                    
                logger.info("Successfully updated aliases")
                return True
            else:
                logger.error("Required columns 'alias' and 'value' not found in Excel file")
                return False
                
        except Exception as e:
            logger.error(f"Error updating aliases: {str(e)}")
            return False

    def get_quote_master_files(self):
        '''
        Returns a list of all Quote Master files in the excel directory
        '''
        try:
            files = [
                file for file in os.listdir(self.excel_path)
                if os.path.isfile(os.path.join(self.excel_path, file)) 
                and file.endswith('.xlsx')
                and file != "parts_sandbox.xlsx"
            ]
            logger.info(f"Found {len(files)} quote master files")
            return files
        except Exception as e:
            logger.error(f"Error listing files: {str(e)}")
            return []

    def search_parts(self, search_term):
        '''
        Searches for parts in the database matching the search term
        '''
        try:
            logger.info(f"Searching for parts with term: {search_term}")
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                # Search in aliases table
                cursor.execute('''
                    SELECT * FROM aliases 
                    WHERE alias LIKE ? OR value LIKE ?
                ''', (f'%{search_term}%', f'%{search_term}%'))
                
                results = cursor.fetchall()
                logger.info(f"Found {len(results)} matching parts")
                
                return [{'alias': row[0], 'value': row[1]} for row in results]
        except Exception as e:
            logger.error(f"Error searching parts: {str(e)}")
            return []

    def get_eau_forecast(self, part_number):
        '''
        Gets EAU forecast for a specific part number from the database
        '''
        try:
            logger.info(f'Loading workbook at {file_path}')
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook['Master Part List']

            #Make a dataframe of the sheet
            df = pd.DataFrame(sheet.values)
            logger.debug(df.head())
        except FileNotFoundError:
            raise FileNotFoundError(f"File {file_path} not found.")

    def refresh_all_files(self):
        """
        Refreshes the database with all Quote Master files.
        Returns True if successful, False if there were any non-critical errors.
        Raises Exception for critical errors.
        """
        try:
            logger.info("Starting database refresh")
            excel_folder = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'excels'))
            sandbox_path = os.path.join(excel_folder, 'parts_sandbox.xlsx')
            
            # Get list of Quote Master files
            files = [
                os.path.join(excel_folder, f) 
                for f in os.listdir(excel_folder) 
                if f.endswith('.xlsx') and f != "parts_sandbox.xlsx"
            ]
            
            if not files:
                logger.warning("No Quote Master files found")
                return False
                
            # Load or create aliases sheet
            try:
                existing_aliases = pd.read_excel(sandbox_path, sheet_name='aliases')
                logger.debug("Loaded existing aliases sheet")
            except (FileNotFoundError, KeyError):
                existing_aliases = pd.DataFrame(columns=['alias', 'value'])
                logger.info("Created new aliases sheet")
            
            had_warnings = False
            all_aliases = []
            
            # Process each file
            for file in files:
                try:
                    logger.debug(f"Processing file: {file}")
                    qm_file = QMFile(file)
                    master_sheet = qm_file.load_master_sheet()
                    df = pd.DataFrame(master_sheet.values)
                    
                    # Skip header row and get alias/value columns
                    df.columns = df.iloc[0]
                    df = df.iloc[1:]
                    
                    if 'alias' in df.columns and 'value' in df.columns:
                        aliases_from_file = df[['alias', 'value']].copy()
                        all_aliases.append(aliases_from_file)
                        logger.debug(f"Successfully processed {len(aliases_from_file)} aliases from {file}")
                    else:
                        logger.warning(f"File {file} missing required columns")
                        had_warnings = True
                        
                except Exception as e:
                    logger.error(f"Error processing {file}: {str(e)}", exc_info=True)
                    had_warnings = True
                    continue
                finally:
                    if hasattr(qm_file, 'workbook'):
                        qm_file.workbook.close()
            
            if all_aliases:
                # Combine all new aliases
                new_aliases = pd.concat(all_aliases, ignore_index=True)
                new_aliases = new_aliases.drop_duplicates(subset=['alias'], keep='first')
                
                # Merge with existing aliases, keeping existing ones in case of conflict
                combined_aliases = pd.concat([existing_aliases, new_aliases], ignore_index=True)
                combined_aliases = combined_aliases.drop_duplicates(subset=['alias'], keep='first')
                
                # Save back to parts_sandbox.xlsx
                with pd.ExcelWriter(sandbox_path, engine='openpyxl', mode='a') as writer:
                    combined_aliases.to_excel(writer, sheet_name='aliases', index=False)
                
                logger.info(f"Successfully saved {len(combined_aliases)} aliases to database")
            
            return not had_warnings
            
        except Exception as e:
            logger.error("Critical error during refresh", exc_info=True)
            raise Exception(f"Error preparing master file: {str(e)}")
        finally:
            if hasattr(self, 'workbook'):
                self.workbook.close()