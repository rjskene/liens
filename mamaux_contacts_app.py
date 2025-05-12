import os
from abc import ABC, abstractmethod

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from liens import (
    create_job_contacts_file, 
    load_job_contacts_files,
    load_existing_jobs,
    load_invoices,
    filter_mamaux_contacts_for_existing_jobs,
    convert_job_list_to_mamaux_format,
    filter_job_contacts_for_invoice_file,
    append_missing_jobs_to_mamaux_contacts
)

class ContactFileFrame(ttk.Frame):
    def __init__(self, parent, available_companies, on_delete=None):
        super().__init__(parent)
        self.available_companies = available_companies
        self.on_delete = on_delete
        
        self.company_var = tk.StringVar()
        self.file_path = tk.StringVar()
        
        self.create_widgets()
        
    def create_widgets(self):
        # Company dropdown
        ttk.Label(self, text="Contact File:").grid(row=0, column=0, sticky=tk.W)
        self.company_dropdown = ttk.Combobox(
            self, 
            textvariable=self.company_var,
            values=list(self.available_companies),
            state='readonly',
            width=10
        )
        self.company_dropdown.grid(row=0, column=1, padx=5)
        
        # File path entry
        self.file_entry = ttk.Entry(
            self,
            textvariable=self.file_path,
            width=50
        )
        self.file_entry.grid(row=0, column=2, padx=5)
        
        # Browse button
        self.browse_btn = ttk.Button(
            self,
            text="Browse",
            command=self.browse_file
        )
        self.browse_btn.grid(row=0, column=3, padx=5)
        
        # Delete button
        self.delete_btn = ttk.Button(
            self,
            text="-",
            command=self.delete_frame,
            width=3
        )
        self.delete_btn.grid(row=0, column=4, padx=5)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
        )
        if filename:
            self.file_path.set(filename)
            
    def delete_frame(self):
        if self.on_delete:
            self.on_delete(self)
        self.destroy()
        
    def get_data(self):
        return self.company_var.get(), self.file_path.get()

class FileUploadBase(ABC):
    """Abstract base class for file upload functionality
    
    ### NOTE ###
    + This class is used to ensure that the subclasses have the required attributes
    + This class is not meant to be instantiated directly
    + Subclasses must define the COMPANIES set
    + Subclasses must define the _contact_frames property
    + subclasses must define file_paths and upload_file_keys properties
    """
    
    @property
    @abstractmethod
    def COMPANIES(self):
        raise NotImplementedError("Subclasses must define COMPANIES set")
    
    @property
    @abstractmethod
    def contact_frames(self):
        raise NotImplementedError("Subclasses must define _contact_frames list")

    @property
    @abstractmethod
    def file_paths(self):
        raise NotImplementedError("Subclasses must define file_paths dictionary")
    
    @property
    @abstractmethod
    def upload_file_keys(self):
        raise NotImplementedError("Subclasses must define upload_file_keys list")

class FileUploadMixin:
    """Mixin class to handle contact frames functionality"""
    
    @property
    def COMPANIES(self):
        return self._COMPANIES
    
    @COMPANIES.setter
    def COMPANIES(self, value):
        self._COMPANIES = value
        self.available_companies = value.copy()

    @property
    def contact_frames(self):
        return self._contact_frames
    
    @contact_frames.setter 
    def contact_frames(self, value):
        self._contact_frames = value
        self.update_button_states()
        
    def browse_file(self, file_type):
        filename = filedialog.askopenfilename(
            filetypes=[("All Files", "*.*"), ("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("Excel Macro Files", "*.xlsm")]
        )
        if filename:
            self.file_paths[file_type].set(filename)

    def create_file_upload_section(self, section, row=1):
        """Create file upload section with labels, entries and browse buttons"""
        for label, key in self.upload_file_keys:
            ttk.Label(section, text=label).grid(row=row, column=0, sticky=tk.W)
            ttk.Entry(section, textvariable=self.file_paths[key], width=50).grid(row=row, column=1, padx=5)
            ttk.Button(section, text="Browse", command=lambda k=key: self.browse_file(k)).grid(row=row, column=2)
            row += 1
            
    def update_button_states(self):
        """Update states of buttons based on contact frames"""
        for frame in self.contact_frames:
            frame.delete_btn['state'] = 'normal' if len(self.contact_frames) > 1 else 'disabled'
            
    def update_frame_positions(self):
        """Update grid positions of contact frames"""
        for i, frame in enumerate(self.contact_frames):
            frame.grid(row=i, column=0, pady=2, sticky='w')

    def add_contact_frame(self):
        if not self.available_companies:
            messagebox.showwarning("Warning", "All company types are already added!")
            return
        
        frame = ContactFileFrame(
            self.contacts_container,
            self.available_companies,
            on_delete=self.on_contact_frame_delete
        )
        # Add button
        frame.add_btn = ttk.Button(
            frame,
            text="+",
            command=self.add_contact_frame,
            width=3
        )
        frame.add_btn.grid(row=0, column=5, padx=1)
        
        # Add frame to list and update grid positions
        self.contact_frames.append(frame)
        self.update_frame_positions()
        
        # Update available companies when a selection is made
        frame.company_dropdown.bind('<<ComboboxSelected>>', 
            lambda e, f=frame: self.update_available_companies(f))

        self.update_button_states()

    def on_contact_frame_delete(self, frame):
        self.contact_frames.remove(frame)
        # Add the company back to available companies
        if frame.company_var.get():
            self.available_companies.add(frame.company_var.get())
        self.update_available_companies(frame)
        self.update_button_states()
        
    def update_available_companies(self, frame):
        # Remove the newly selected company from available_companies
        selected = frame.company_var.get()
        if selected:
            self.available_companies = self.COMPANIES - {f.company_var.get() for f in self.contact_frames if f.company_var.get()}
            
            # Update dropdown values for frames without selection
            for f in self.contact_frames:
                if not f.company_var.get():
                    f.company_dropdown['values'] = list(self.available_companies)

class MamauxContactsApp(FileUploadMixin, FileUploadBase):
    def __init__(self, root):
        self.root = root
        self.root.title("Mamaux Contacts Manager")
        self.root.geometry("700x250")
        
        self._contact_frames = []
        self.COMPANIES = {'HTS', 'DXS'}
        self._file_paths = {
            'invoices_file': tk.StringVar(),
            'liens_template_file': tk.StringVar(),
        }
        self._upload_file_keys = [
            ('Invoices File:', 'invoices_file'),
            ('Liens Template File:', 'liens_template_file')
        ]
        self.create_widgets()
    
    @property
    def file_paths(self):
        return self._file_paths
    
    @property
    def upload_file_keys(self):
        return self._upload_file_keys

    def create_widgets(self):
        # Create main frame with scrollbar
        self.canvas = tk.Canvas(self.root)
        self.main_frame = ttk.Frame(self.canvas)        
        self.canvas.pack(side="left", fill="both", expand=True)        
        # Create window in canvas
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")

        # Bind Crtl + W to close the app
        self.root.bind("<Control-w>", self.close_app)
        
       # FRAMES
        upload_frame = ttk.Frame(self.main_frame) # Frame to hold both sections side by side
        upload_frame.grid(row=0, column=0, columnspan=2, pady=5, sticky='W')
        ttk.Label(upload_frame, text="Generate Mamaux Contacts", font=('Helvetica', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=10)

        upload_section = ttk.Frame(upload_frame)
        upload_section.grid_configure(sticky='W')
        upload_section.grid(row=2, column=0, padx=5)

        self.contacts_container = ttk.Frame(upload_frame)
        self.contacts_container.grid_configure(sticky='W')
        self.contacts_container.grid(row=4, column=0, pady=20)

        self.create_file_upload_section(upload_section)
        self.add_contact_frame() # Initialize with one contact frame

        # Process button
        ttk.Button(
            upload_frame, 
            text="Create Lien Contacts for Mamaux", 
            command=self.generate_mamaux_contacts,
            width=30,
        ).grid(row=6, column=0, columnspan=3, pady=20)
         
    def close_app(self, event=None):
        self.root.destroy()

    def generate_mamaux_contacts(self):
        try:
            # Get contacts files dictionary
            contacts_files = {}
            for frame in self.contact_frames:
                company, file_path = frame.get_data()
                if not company or not file_path:
                    messagebox.showerror("Error", "Please complete all contact file entries")
                    return
                contacts_files[company] = file_path

            # Load all data
            contacts_dfs = load_job_contacts_files(contacts_files)
            df_existing_jobs = load_existing_jobs(self.file_paths['liens_template_file'].get())
            df_invs = load_invoices(self.file_paths['invoices_file'].get())

            # Process the data
            df_job_contacts = create_job_contacts_file(contacts_dfs)
            df_job_contacts = filter_job_contacts_for_invoice_file(df_job_contacts, df_invs)
            df_mamaux_contacts = convert_job_list_to_mamaux_format(df_job_contacts)
            df_mamaux_contacts = filter_mamaux_contacts_for_existing_jobs(df_mamaux_contacts, df_existing_jobs)
            self.df_mamaux_contacts = append_missing_jobs_to_mamaux_contacts(df_mamaux_contacts, df_job_contacts, df_invs)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while generating contacts:\n{str(e)}")
    
        try:
            # Open file dialog to get save location
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")]
            )
            
            if filename:
                # Save missing jobs to CSV
                self.df_mamaux_contacts.to_csv(filename, index=False)
                messagebox.showinfo("Success", "Mamaux contacts file saved successfully!")

                # Open the file
                os.startfile(filename)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving the file:\n{str(e)}")

def main():
    root = tk.Tk()
    app = MamauxContactsApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()