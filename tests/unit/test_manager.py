"""
Unit tests for the ApplicationManager class.
Tests database operations and file handling functionality.
"""
import pytest
import os
import sqlite3
import pandas as pd
from unittest.mock import Mock, patch, PropertyMock
from src.manager import ApplicationManager

@pytest.fixture(autouse=True)
def app_manager():
    """Fixture to create a test ApplicationManager instance with clean database"""
    manager = ApplicationManager()
    manager.db_path = ':memory:'  # Use in-memory database for testing
    manager.excel_path = 'test_excels'
    # Initialize the database
    with manager.get_connection() as conn:
        manager._init_db(conn)
    return manager

def test_check_for_db(app_manager):
    """Test database initialization"""
    with app_manager.get_connection() as conn:
        cursor = conn.cursor()
        
        # Check if tables exist
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='aliases'")
        assert cursor.fetchone() is not None, "Aliases table was not created"
        
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='quote_masters'")
        assert cursor.fetchone() is not None, "Quote masters table was not created"

@pytest.mark.parametrize("search_term,expected_count", [
    ("test_part", 1),
    ("nonexistent", 0),
])
def test_search_parts(app_manager, search_term, expected_count):
    """Test parts search functionality"""
    # Insert test data
    with app_manager.get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM aliases")  # Clean existing data
        cursor.execute("INSERT INTO aliases (alias, value) VALUES (?, ?)", 
                      ("test_part", "TEST123"))
    
    # Test the search functionality
    results = app_manager.search_parts(search_term)
    assert len(results) == expected_count
    if expected_count > 0:
        assert results[0]['alias'] == "test_part"
        assert results[0]['value'] == "TEST123"

@patch('pandas.read_excel')
def test_update_alias(mock_read_excel, app_manager):
    """Test alias update functionality"""
    # Create test DataFrame
    test_df = pd.DataFrame({
        'alias': ['test_alias1', 'test_alias2'],
        'value': ['TEST123', 'TEST456']
    })
    mock_read_excel.return_value = test_df
    
    # Test the update functionality
    result = app_manager.update_alias("dummy_file.xlsx")
    assert result is True, "Alias update should succeed with valid file"
    
    # Verify the data was properly inserted
    with app_manager.get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM aliases")
        count = cursor.fetchone()[0]
        assert count == 2, "Expected 2 aliases to be inserted"

def test_get_quote_master_files(app_manager, tmp_path):
    """Test quote master file listing"""
    # Create temporary test files
    test_dir = tmp_path / "test_excels"
    test_dir.mkdir()
    (test_dir / "test_quote.xlsx").touch()
    (test_dir / "parts_sandbox.xlsx").touch()
    
    with patch.object(app_manager, 'excel_path', str(test_dir)):
        files = app_manager.get_quote_master_files()
        assert len(files) == 1, "Should only return non-parts_sandbox xlsx files"
        assert "test_quote.xlsx" in files