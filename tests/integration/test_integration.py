"""
Integration tests for the Parts Sandbox application.
Tests interactions between different components and end-to-end functionality.
"""
import pytest
import os
import sqlite3
import pandas as pd
from src.manager import ApplicationManager
from src.QMFile import QMFile
import tempfile
import shutil

@pytest.fixture
def test_env():
    """Create a temporary test environment with necessary files and directories"""
    # Create temporary directory
    temp_dir = tempfile.mkdtemp()
    excel_dir = os.path.join(temp_dir, 'excels')
    db_dir = os.path.join(temp_dir, 'database')
    os.makedirs(excel_dir)
    os.makedirs(db_dir)
    
    # Create a test Quote Master file with proper sheet name and structure
    test_df = pd.DataFrame({
        'Part Number': ['TEST001', 'TEST002'],
        'Description': ['Test Part 1', 'Test Part 2'],
        'EAU': [100, 200],
        'alias': ['test_part_1', 'test_part_2'],
        'value': ['TEST001', 'TEST002']
    })
    test_excel_path = os.path.join(excel_dir, 'test_quote_master.xlsx')
    with pd.ExcelWriter(test_excel_path, engine='openpyxl') as writer:
        test_df.to_excel(writer, sheet_name='Master Part List', index=False)
    
    # Create a test parts_sandbox.xlsx file
    sandbox_df = pd.DataFrame({
        'Part Number': ['TEST001', 'TEST002'],
        'Description': ['Test Part 1', 'Test Part 2'],
        'EAU': [100, 200]
    })
    sandbox_path = os.path.join(excel_dir, 'parts_sandbox.xlsx')
    with pd.ExcelWriter(sandbox_path, engine='openpyxl') as writer:
        sandbox_df.to_excel(writer, index=False)
    
    yield {
        'temp_dir': temp_dir,
        'excel_dir': excel_dir,
        'db_dir': db_dir,
        'test_excel_path': test_excel_path,
        'sandbox_path': sandbox_path
    }
    
    # Cleanup
    shutil.rmtree(temp_dir)

def test_end_to_end_quote_master_flow(test_env):
    """Test the complete flow of loading and processing a Quote Master file"""
    # Initialize ApplicationManager with test environment
    app_manager = ApplicationManager()
    app_manager.db_path = os.path.join(test_env['db_dir'], 'test.db')
    app_manager.excel_path = test_env['excel_dir']
    app_manager.start()
    
    # Test quote master file listing
    files = app_manager.get_quote_master_files()
    assert 'test_quote_master.xlsx' in files
    
    # Test loading quote master file
    qm_file = QMFile(test_env['test_excel_path'])
    assert qm_file is not None
    assert 'parts' in qm_file.data
    assert len(qm_file.data['parts']) == 2
    
    # Test updating aliases from quote master
    success = app_manager.update_alias(test_env['test_excel_path'])
    assert success is True
    
    # Test searching for parts
    results = app_manager.search_parts('TEST')
    assert len(results) > 0
    
    # Test EAU forecast
    forecast = app_manager.get_eau_forecast('TEST001')
    assert forecast['eau'] == 100

def test_database_persistence(test_env):
    """Test that data persists correctly in the database"""
    app_manager = ApplicationManager()
    app_manager.db_path = os.path.join(test_env['db_dir'], 'test.db')
    app_manager.excel_path = test_env['excel_dir']
    app_manager.start()
    
    # Update aliases
    app_manager.update_alias(test_env['test_excel_path'])
    
    # Create new instance to test persistence
    new_manager = ApplicationManager()
    new_manager.db_path = app_manager.db_path
    new_manager.excel_path = app_manager.excel_path
    new_manager.start()  # Make sure tables are initialized
    
    # Verify data persists
    results = new_manager.search_parts('TEST')
    assert len(results) > 0