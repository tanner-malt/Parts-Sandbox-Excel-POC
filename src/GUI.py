'''
Tkinter GUI for the application - Client Mode
This module creates a graphical user interface (GUI) using Tkinter for the application.
'''

import tkinter as tk
from tkinter import messagebox  # Import messagebox for error dialogs
import mainfuncs as mf # These are localized functions I wanted seperate to maintain fascade of being clean
import os  # Import os to handle file opening
import logging

logger = logging.getLogger(__name__)
app_manager = None

def start_application(manager):
    """
    Initializes and starts the main application window.
    """
    global app_manager
    app_manager = manager
    logger.info("Initializing main application window")
    
    # Create the main application window
    main_window = tk.Tk()
    main_window.title("Parts Sandbox Manager - Client")
    main_window.geometry("800x600")

    # Create a label to display a welcome message
    welcome_label = tk.Label(main_window, text="Welcome to the Parts Sandbox Manager!", font=("Arial", 16))
    welcome_label.pack(pady=20)

    # Create a button to list all Quote Master Files
    list_button = tk.Button(main_window, text="List Quote Master Files", command=list_files, font=("Arial", 12))
    list_button.pack(pady=10) 

    list_button = tk.Button(main_window, text="Analysis Mode", command=mf.analysis_loop, font=("Arial", 12))
    list_button.pack(pady=10) 

    list_button = tk.Button(main_window, text="EAU Forecast", command=lambda: print("dummy"), font=("Arial", 12))
    list_button = tk.Button(main_window, text="EAU Forecast", command=lambda: print("dummy"), font=("Arial", 12))
    list_button.pack(pady=10)

    list_button = tk.Button(main_window, text="Search Parts", command=lambda: print("dummy"), font=("Arial", 12))
    list_button.pack(pady=10) 

    # Create a button to exit the application
    exit_button = tk.Button(main_window, text="Exit", command=main_window.quit, font=("Arial", 12))
    exit_button.pack(pady=10)

    logger.info("Main window initialized, starting event loop")
    # Start the Tkinter event loop
    main_window.mainloop()

def list_files():
    '''
    This function will list all the files in the "excel" folder in the GUI.
    '''
    logger.info("Opening Quote Master Files window")
    excel_folder = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'excels'))

    try:
        quote_master_files = [
            os.path.join(excel_folder, file)
            for file in os.listdir(excel_folder)
            if os.path.isfile(os.path.join(excel_folder, file)) and file != "parts_sandbox.xlsx"
        ]
        logger.debug(f"Found {len(quote_master_files)} Quote Master files")
    except FileNotFoundError:
        logger.error(f"Excel folder not found: {excel_folder}")
        messagebox.showerror("Error", f"The folder '{excel_folder}' does not exist.")
        return

    file_window = tk.Toplevel()
    file_window.title("Quote Master Files")
    file_window.geometry("600x600")

    file_label = tk.Label(file_window, text="Quote Master Files:", font=("Arial", 12))
    file_label.pack(pady=10)

    # Create a Listbox to display the files
    file_listbox = tk.Listbox(file_window, font=("Arial", 10), width=50, height=15, selectmode=tk.SINGLE)
    file_listbox.pack(pady=10)

    for file in quote_master_files:
        file_listbox.insert(tk.END, file)

    # Define a function to open the selected file
    def open_file(event=None):
        if not file_listbox.curselection():
            logger.warning("No file selected for opening")
            messagebox.showerror("Error", "Please select a file first")
            return
            
        selected_file = file_listbox.get(file_listbox.curselection())
        if os.path.exists(selected_file):
            logger.info(f"Opening file: {selected_file}")
            os.startfile(selected_file)
        else:
            logger.error(f"File not found: {selected_file}")
            messagebox.showerror("Error", f"File not found: {selected_file}")

    def refresh_database():
        try:
            logger.info("Starting database refresh")
            # Use the global app_manager instance
            success = app_manager.refresh_all_files()
            if success:
                logger.info("Database refresh completed successfully")
                messagebox.showinfo("Success", "Database refreshed successfully with all Quote Master files")
            else:
                logger.warning("Database refresh completed with warnings")
                messagebox.showwarning("Warning", "Database refresh completed with some warnings. Check the log for details.")
        except Exception as e:
            logger.error(f"Database refresh failed: {str(e)}", exc_info=True)
            messagebox.showerror("Error", f"Failed to refresh database: {str(e)}")

    # Create buttons frame
    button_frame = tk.Frame(file_window)
    button_frame.pack(pady=10)

    refresh_button = tk.Button(button_frame, text="Refresh All", command=refresh_database, font=("Arial", 10))
    refresh_button.pack(side=tk.LEFT, padx=5)

    close_button = tk.Button(button_frame, text="Close", command=file_window.destroy, font=("Arial", 10))
    close_button.pack(side=tk.LEFT, padx=5)

    # Bind double-click event to the Listbox for opening files
    file_listbox.bind("<Double-1>", open_file)

def make_analysis_window():
    '''
    This function will create the analysis window, and any other items needed
    '''
    analysis_window = tk.Toplevel()
    analysis_window.title("Analysis Mode")
    analysis_window.geometry("600x600")

    # Add alias display or other analysis features here
    alias_label = tk.Label(analysis_window, text="Available Aliases:", font=("Arial", 12))
    alias_label.pack(pady=10)

    alias_listbox = tk.Listbox(analysis_window, font=("Arial", 10), width=50, height=15)
    alias_listbox.pack(pady=10)

    for alias in aliases:
        alias_listbox.insert(tk.END, f"{alias['alias']} -> {alias['value']}")

    close_button = tk.Button(analysis_window, text="Close", command=analysis_window.destroy, font=("Arial", 10))
    close_button.pack(pady=10)