'''
Tkinter GUI for the application
This module creates a graphical user interface (GUI) using Tkinter for the application.
'''

import tkinter as tk
from tkinter import messagebox  # Import messagebox for error dialogs
import mainfuncs as mf
import os  # Import os to handle file opening

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
    list_button = tk.Button(main_window, text="List Quote Master Files", command=list_files, font=("Arial", 12))
    list_button.pack(pady=10) 

    list_button = tk.Button(main_window, text="Analyze Files", command=print("dummy"), font=("Arial", 12))
    list_button.pack(pady=10) 

    list_button = tk.Button(main_window, text="EAU Forecast", command=print("dummy"), font=("Arial", 12))
    list_button.pack(pady=10)

    list_button = tk.Button(main_window, text="Search Parts", command=print("dummy"), font=("Arial", 12))
    list_button.pack(pady=10) 

    # Create a button to exit the application
    exit_button = tk.Button(main_window, text="Exit", command=main_window.quit, font=("Arial", 12))
    exit_button.pack(pady=10)

    # Start the Tkinter event loop
    main_window.mainloop()


def list_files():
    '''
    This function will list all the files in the "excel" folder in the GUI.
    It will be called when the user clicks the "List Quote Master Files" button.
    '''
    # Define the path to the "excel" folder relative to the project root
    excel_folder = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'excels'))

    # Get the list of files in the "excel" folder
    try:
        quote_master_files = [
            os.path.join(excel_folder, file)
            for file in os.listdir(excel_folder)
            if os.path.isfile(os.path.join(excel_folder, file)) and file != "parts_sandbox"
        ]
    except FileNotFoundError:
        messagebox.showerror("Error", f"The folder '{excel_folder}' does not exist.")
        return

    # Create a new window to display the list of files
    file_window = tk.Toplevel()
    file_window.title("Quote Master Files")
    file_window.geometry("600x600")

    # Create a label to display the title
    file_label = tk.Label(file_window, text="Quote Master Files:", font=("Arial", 12))
    file_label.pack(pady=10)

    # Create a Listbox to display the files
    file_listbox = tk.Listbox(file_window, font=("Arial", 10), width=50, height=15)
    file_listbox.pack(pady=10)

    # Insert each file into the Listbox
    for file in quote_master_files:
        file_listbox.insert(tk.END, file)

    # Define a function to open the selected file
    def open_file(event):
        try:
            # Get the selected file
            selected_file = file_listbox.get(file_listbox.curselection())
            if os.path.exists(selected_file):
                os.startfile(selected_file)
            else:
                messagebox.showerror("Error", f"File not found: {selected_file}")
        except IndexError:
            messagebox.showerror("Error", "No file selected.")

    #Dummy function to update alias, needs to be implemented.
    def show_context_menu(event):
        # Create a context menu
        file_listbox.selection_clear(0, tk.END)
        file_listbox.selection_set(file_listbox.nearest(event.y))
        selected_index = file_listbox.curselection()
        if selected_index:
            context_menu = tk.Menu(file_window, tearoff=0)
            context_menu.add_command(label="Open", command=open_file)
            context_menu.add_command(label="Update Alias",command = print("Update Alias"))
            context_menu.post(event.x_root, event.y_root)


    # Bind double-click event to the Listbox
    file_listbox.bind("<Double-1>", open_file)
    file_listbox.bind("<Button-3>", show_context_menu)  # Right-click to show context menu

    # Add a close button
    close_button = tk.Button(file_window, text="Close", command=file_window.destroy, font=("Arial", 10))
    close_button.pack(pady=10)