'''
Tkinter GUI for the application
This module creates a graphical user interface (GUI) using Tkinter for the application.
'''

import tkinter as tk
import mainfuncs as mf

def start_application():
    """
    Initializes and starts the main application window.
    """
    # Create the main application window
    main_window = tk.Tk()
    main_window.title("Parts Sandbox Manager")
    main_window.geometry("800x600")

    # Create a label to display a welcome message
    welcome_label = tk.Label(main_window, text="Welcome to the Parts Sandbox Manager!", font=("Arial", 16))
    welcome_label.pack(pady=20)

    # Create a button to list all Quote Master Files
    list_button = tk.Button(main_window, text="List Quote Master Files", command=mf.list_quote_master_files, font=("Arial", 12))
    list_button.pack(pady=10) 

    # Create a button to exit the application
    exit_button = tk.Button(main_window, text="Exit", command=main_window.quit, font=("Arial", 12))
    exit_button.pack(pady=10)


    # Start the Tkinter event loop
    main_window.mainloop()