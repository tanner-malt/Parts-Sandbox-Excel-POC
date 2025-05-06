'''
Tkinter GUI for the application - Client Mode
This module creates a graphical user interface (GUI) using Tkinter for the application.
'''

import tkinter as tk
from tkinter import messagebox
import os
import requests
from urllib.parse import urljoin

class Config:
    SERVER_URL = 'http://localhost:5000'  # Can be changed for different network setups

def start_application():
    """
    Initializes and starts the main application window.
    """
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

    list_button = tk.Button(main_window, text="Analysis Mode", command=lambda: show_analysis_window(main_window), font=("Arial", 12))
    list_button.pack(pady=10) 

    list_button = tk.Button(main_window, text="EAU Forecast", command=lambda: print("dummy"), font=("Arial", 12))
    list_button.pack(pady=10)

    list_button = tk.Button(main_window, text="Search Parts", command=lambda: print("dummy"), font=("Arial", 12))
    list_button.pack(pady=10) 

    # Create a button to exit the application
    exit_button = tk.Button(main_window, text="Exit", command=main_window.quit, font=("Arial", 12))
    exit_button.pack(pady=10)

    # Start the Tkinter event loop
    main_window.mainloop()

def list_files():
    '''
    This function will list all the files in the "excel" folder in the GUI.
    '''
    excel_folder = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'excels'))

    try:
        quote_master_files = [
            os.path.join(excel_folder, file)
            for file in os.listdir(excel_folder)
            if os.path.isfile(os.path.join(excel_folder, file)) and file != "parts_sandbox.db"
        ]
    except FileNotFoundError:
        messagebox.showerror("Error", f"The folder '{excel_folder}' does not exist.")
        return

    file_window = tk.Toplevel()
    file_window.title("Quote Master Files")
    file_window.geometry("600x600")

    file_label = tk.Label(file_window, text="Quote Master Files:", font=("Arial", 12))
    file_label.pack(pady=10)

    file_listbox = tk.Listbox(file_window, font=("Arial", 10), width=50, height=15)
    file_listbox.pack(pady=10)

    for file in quote_master_files:
        file_listbox.insert(tk.END, file)

    def open_file(event):
        try:
            selected_file = file_listbox.get(file_listbox.curselection())
            if os.path.exists(selected_file):
                os.startfile(selected_file)
            else:
                messagebox.showerror("Error", f"File not found: {selected_file}")
        except IndexError:
            messagebox.showerror("Error", "No file selected.")

    def update_alias():
        try:
            selected_file = file_listbox.get(file_listbox.curselection())
            if not os.path.exists(selected_file):
                messagebox.showerror("Error", f"File not found: {selected_file}")
                return

            # Send file to server
            with open(selected_file, 'rb') as f:
                files = {'file': f}
                response = requests.post(urljoin(Config.SERVER_URL, 'api/update_alias'), files=files)
            
            if response.status_code == 200 and response.json().get('success'):
                messagebox.showinfo("Success", "Alias updated successfully")
            else:
                messagebox.showerror("Error", "Failed to update alias")
        except IndexError:
            messagebox.showerror("Error", "No file selected.")
        except requests.RequestException as e:
            messagebox.showerror("Error", f"Server communication error: {str(e)}")

    def show_context_menu(event):
        file_listbox.selection_clear(0, tk.END)
        file_listbox.selection_set(file_listbox.nearest(event.y))
        selected_index = file_listbox.curselection()
        if selected_index:
            context_menu = tk.Menu(file_window, tearoff=0)
            context_menu.add_command(label="Open", command=lambda: open_file(None))
            context_menu.add_command(label="Update Alias", command=update_alias)
            context_menu.post(event.x_root, event.y_root)

    file_listbox.bind("<Double-1>", open_file)
    file_listbox.bind("<Button-3>", show_context_menu)

    close_button = tk.Button(file_window, text="Close", command=file_window.destroy, font=("Arial", 10))
    close_button.pack(pady=10)

def show_analysis_window(parent):
    '''
    Creates and shows the analysis window
    '''
    try:
        # Fetch aliases from server
        response = requests.get(urljoin(Config.SERVER_URL, 'api/aliases'))
        if response.status_code == 200:
            aliases = response.json()
        else:
            messagebox.showerror("Error", "Failed to fetch aliases from server")
            return
    except requests.RequestException as e:
        messagebox.showerror("Error", f"Server communication error: {str(e)}")
        return

    analysis_window = tk.Toplevel(parent)
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