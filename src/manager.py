'''
This file is the Manager for the Parts Sandbox application.
It is responsible for managing the overall flow of the application, including database operations and request handling.
'''
import sqlite3
import pandas as pd
from flask import Flask, jsonify, request
import os
import logging
from contextlib import contextmanager

app = Flask(__name__)
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
        if ':memory:' not in self.db_path and not os.path.exists(os.path.dirname(self.db_path)):
            logger.info(f"Creating database directory at {os.path.dirname(self.db_path)}")
            os.makedirs(os.path.dirname(self.db_path))
            
        with self.get_connection() as conn:
            self._init_db(conn)
            logger.info("Database initialization complete")

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
            logger.info(f"Getting EAU forecast for part: {part_number}")
            # For now, we'll read from the parts_sandbox.xlsx file
            df = pd.read_excel(os.path.join(self.excel_path, 'parts_sandbox.xlsx'))
            forecast = df[df['Part Number'] == part_number]['EAU'].iloc[0] if 'EAU' in df.columns else 0
            return {'part_number': part_number, 'eau': float(forecast)}
        except Exception as e:
            logger.error(f"Error getting EAU forecast: {str(e)}")
            return {'part_number': part_number, 'eau': 0}

    def __del__(self):
        """Close the database connection when the object is destroyed"""
        if self._conn is not None:
            self._conn.close()
            self._conn = None

# Flask routes
@app.route('/api/aliases', methods=['GET'])
def get_aliases():
    manager = ApplicationManager()
    with manager.get_connection() as conn:
        df = pd.read_sql_query("SELECT * FROM aliases", conn)
    return jsonify(df.to_dict('records'))

@app.route('/api/update_alias', methods=['POST'])
def update_alias_endpoint():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
        
    manager = ApplicationManager()
    success = manager.update_alias(file)
    return jsonify({'success': success})

@app.route('/api/quote_master_files', methods=['GET'])
def get_quote_master_files():
    manager = ApplicationManager()
    files = manager.get_quote_master_files()
    return jsonify({'files': files})

@app.route('/api/search_parts', methods=['GET'])
def search_parts():
    search_term = request.args.get('term', '')
    if not search_term:
        return jsonify({'error': 'No search term provided'}), 400
    
    manager = ApplicationManager()
    results = manager.search_parts(search_term)
    return jsonify({'results': results})

@app.route('/api/eau_forecast/<part_number>', methods=['GET'])
def get_eau_forecast(part_number):
    if not part_number:
        return jsonify({'error': 'No part number provided'}), 400
    
    manager = ApplicationManager()
    forecast = manager.get_eau_forecast(part_number)
    return jsonify(forecast)

def start_server():
    '''
    Starts the Flask server
    '''
    app.run(host='0.0.0.0', port=5000)