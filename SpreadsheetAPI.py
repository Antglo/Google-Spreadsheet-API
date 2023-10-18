#Application Created to make submitting warranty requests easier.
#Antonio Taboas 10/17/2023
#Warranty Parts Request through oAuth
import sys
import os
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk
import gspread
from oauth2client.service_account import ServiceAccountCredentials

#Use credentials to create a client to interact with the Google Drive API
#Make a function that appends to sheet
class warranty_parts_request_form:
    def __init__(self, root):
        self.root = root
        self.root.title("Warranty Parts GUI")
        self.root.geometry("550x350") #Set the default window size
        self.root.configure(bg="lightblue")
        
        #Style the tkinter window
        style = ttk.Style()
        style.configure("TLabel", background="lightblue", font=("Arial", 12))
        style.configure("Tbutton", font=("Arial", 12), padding=5)
        style.configure("TEntry", padding=5, font=("Arial", 12))
        
        #Get the current datetime
        current_date = datetime.now().strftime("%m/%d/%Y %H:%M:%S")
        
        #Create date field
        date_field = ttk.Label(self.root, text='Enter Todays date:')
        date_field.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        self.entry_date = ttk.Entry(self.root)
        self.entry_date.grid(row=0, column=2, padx=10, pady=10)
        
        #set the default date to datetime
        self.entry_date.insert(0, current_date)

        #Create portal field
        portal_field = ttk.Label(self.root, text='Enter the portal: (Trafera, Archangel)')
        portal_field.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        self.entry_portal = ttk.Entry(self.root)
        self.entry_portal.grid(row=1, column=2, padx=10, pady=10)
        
        #Create default entry for Trafera
        self.entry_portal.insert(0, "Trafera")

        #Create model field
        model_field = ttk.Label(self.root, text='Enter Chromebook model:')
        model_field.grid(row=2, column=1, padx=10, pady=10, sticky="w")
        #create a dropdown list to choose from
        chromebook_models = ["100e MTK", "100e AST", "100e 2nd gen", "300e MTK", "300e AST"]
        self.entry_model = ttk.Combobox(self.root, values=chromebook_models)
        self.entry_model.grid(row=2, column=2, padx=10, pady=10)

        #Create asset field
        asset_field = ttk.Label(self.root, text='Enter the asset #: (123456)')
        asset_field.grid(row=3, column=1, padx=10, pady=10, sticky="w")
        self.entry_asset = ttk.Entry(self.root)
        self.entry_asset.grid(row=3, column=2, padx=10, pady=10)

        #Create Serial field
        serial_field = ttk.Label(self.root, text='Enter the serial #: (This can be scanned)')
        serial_field.grid(row=4, column=1, padx=10,pady=10, sticky="w")
        self.entry_serial = ttk.Entry(self.root)
        self.entry_serial.grid(row=4, column=2, padx=10, pady=10)

        #Create parts field
        parts_field = ttk.Label(self.root, text='Enter Part type:')
        parts_field.grid(row=5, column=1, padx=10, pady=10, sticky="w")
        #Create dropdown list to choose from
        chromebook_parts = ["Screen", "Battery", "Keyboard", "Motherboard"]
        self.entry_parts = ttk.Combobox(self.root, values=chromebook_parts)
        self.entry_parts.grid(row=5, column=2, padx=10, pady=10)
       

        #Create a submit button
        submit_button = ttk.Button(self.root, text='Insert data into the google sheet', command=self.append_to_sheet)
        submit_button.grid(row=7, pady=5, columnspan=4)
    
    #Make a function that appends to sheet
    def append_to_sheet(self):
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        
        #Determine the path to credentials based on execution
        if getattr(sys, "frozen", False):
            application_path = sys._MEIPASS
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        credentials_path = os.path.join(application_path, "credentials.json")
        
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
        client = gspread.authorize(creds)

        #Find the sheet by name and open the first sheet (front Page)
        sheet = client.open('Warranty Parts Request GUI').sheet1
        
        #With page open find duplicates and raise message if there is a duplicate
        #Be carful column values are (0-based) index start from 0...
        asset_tags = sheet.col_values(3)
        serial_numbers = sheet.col_values(4)
        
        user_input_asset =  self.entry_asset.get()
        user_input_serial = self.entry_serial.get()
        
        #Create the if logic for duplications
        if user_input_asset in asset_tags:
            messagebox.showwarning("Duplicate Entry", "The ASSET_TAG you entered already exists!")
            return #use to exit the function
        if user_input_serial in serial_numbers:
            messagebox.showwarning("Duplicate Entry", "The SERIAL_NUMBER you entered already exists!")

        #append data from GUI fields
        data = [self.entry_date.get(), self.entry_portal.get(), self.entry_model.get(), self.entry_asset.get(), self.entry_serial.get(), self.entry_parts.get()]
        sheet.append_row(data, value_input_option='USER_ENTERED')
        messagebox.showinfo("Success", "Data inserted successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = warranty_parts_request_form(root)
    root.mainloop()