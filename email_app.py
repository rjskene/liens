import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from liens import (
    create_job_contacts_file, 
    connect_to_outlook,
    send_outlook_email,
    load_job_contacts_files,
    load_invoices,
    load_projects,
    load_emails,
    filter_job_contacts_for_invoice_file,
    append_leader_to_job_contacts,
    append_emails_to_job_contacts,
    filter_job_contacts_for_missing_info,
    attach_urls_to_job_contacts,
)
from link_scraper import (
    scrape_for_new_urls,
    PROJECT_LINKS_FILE,
)
from mamaux_contacts_app import FileUploadBase, FileUploadMixin

class EmailApp(FileUploadMixin, FileUploadBase):
    def __init__(self, root):
        self.root = root
        self.root.title("Project Contact Emails")
        self.root.geometry("700x500")
        
        self.COMPANIES = {'HTS', 'DXS', 'ONCO', 'VRFS'}
        self._contact_frames = []
        
        # File paths
        self._file_paths = {
            'users_file': tk.StringVar(),
            'projects_file': tk.StringVar(),
            'invoices_file': tk.StringVar(),
        }
        self._upload_file_keys = [
            ('Users File:', 'users_file'),
            ('Projects File:', 'projects_file'),
            ('Invoices File:', 'invoices_file'),
        ]
        
        self.contact_frames = []

        """FOR TESTING"""
        self._TEST_TO_EMAIL = 'ryan.skene@hts.com'
        self.create_widgets()
    
    @property
    def file_paths(self):
        return self._file_paths
    
    @property
    def upload_file_keys(self):
        return self._upload_file_keys

    def close_app(self, event=None):
        self.root.destroy()

    def create_widgets(self):
        self.canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.main_frame = ttk.Frame(self.canvas)
        
        # Configure canvas and scrollbar
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Create window in canvas
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")

        # Configure canvas scrolling
        self.main_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        # Bind Crtl + W to close the app
        self.root.bind("<Control-w>", self.close_app)

       # FRAMES
        upload_frame = ttk.Frame(self.main_frame) # Frame to hold both sections side by side
        upload_frame.grid(row=0, column=0, columnspan=2, pady=5, sticky='W')
        ttk.Label(upload_frame, text="Find Missing Contacts", font=('Helvetica', 12, 'bold')).grid(row=0, column=0, columnspan=2, pady=10)

        upload_section = ttk.Frame(upload_frame)
        upload_section.grid_configure(sticky='W')
        upload_section.grid(row=2, column=0, padx=5)

        self.contacts_container = ttk.Frame(upload_frame)
        self.contacts_container.grid_configure(sticky='W')
        self.contacts_container.grid(row=4, column=0, pady=20)

        self.create_file_upload_section(upload_section)
        self.add_contact_frame() # Initialize with one contact frame

        # Process buttons
        button_frame = ttk.Frame(upload_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=20)

        ttk.Button(
            button_frame, 
            text="Find Missing Contacts", 
            command=self.find_missing_contacts,
            width=30,
        ).grid(row=0, column=0, padx=5)

        self.download_missing_info_btn = ttk.Button(
            button_frame,
            text="Download Missing Info",
            command=self.download_missing_info, 
            state=tk.DISABLED,
            width=30,
        )
        self.download_missing_info_btn.grid(row=0, column=1, padx=5)

        # Upload existing missing info section
        upload_missing_frame = ttk.Frame(upload_frame)
        upload_missing_frame.grid(row=7, column=0, columnspan=3, pady=20)
        
        ttk.Label(upload_missing_frame, text="Or Upload Existing Missing Info File:").grid(row=0, column=0, padx=5)
        
        self.missing_info_path = tk.StringVar()
        ttk.Entry(upload_missing_frame, textvariable=self.missing_info_path, width=50).grid(row=0, column=1, padx=5)
        
        ttk.Button(
            upload_missing_frame,
            text="Browse",
            command=lambda: self.upload_missing_info_file('missing_info_file'),
            width=10
        ).grid(row=0, column=2, padx=5)

        # Missing Info section
        row = 6
        ttk.Label(self.main_frame, text="Missing Info", font=('Helvetica', 12, 'bold')).grid(row=row, column=0, columnspan=3, pady=(20,10))

        missing_info_frame = ttk.Frame(self.main_frame)
        missing_info_frame.grid(row=row+1, column=0, columnspan=2, pady=10, sticky='nsew')
        
        # Configure grid weights to allow expansion
        missing_info_frame.grid_rowconfigure(0, weight=1)
        missing_info_frame.grid_columnconfigure(0, weight=1)

        text_frame = ttk.Frame(missing_info_frame)
        text_frame.grid(row=0, column=0, columnspan=1, pady=10, sticky='nsew')

        # Add text widget and scrollbar
        self.missing_info_text = tk.Text(text_frame, height=20, width=20, padx=10, pady=10, wrap=tk.WORD)
        missing_scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=self.missing_info_text.yview)
        
        # Configure text widget and scrollbar
        self.missing_info_text.configure(yscrollcommand=missing_scrollbar.set)
        
        # Pack widgets with fill and expand
        self.missing_info_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        missing_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        button_frame = ttk.Frame(missing_info_frame)
        button_frame.grid(row=0, column=1, pady=10, sticky='ns')

        # Send emails button and GO LIVE checkbox
        send_emails_frame = ttk.Frame(button_frame)
        send_emails_frame.grid(row=1, column=1, padx=50, pady=20)
        
        def send_emails_with_confirmation(self):
            if self.go_live_var.get():
                if messagebox.askyesno("Warning", "You are about to send LIVE emails. Are you sure you want to proceed?"):
                    self.send_emails()
            else:
                self.send_emails()
        
        self.send_emails_btn = ttk.Button(
            send_emails_frame,
            text="Send Emails", 
            command=lambda: send_emails_with_confirmation(self),
            state=tk.DISABLED
        )
        self.send_emails_btn.pack(side=tk.LEFT, padx=(0,10))

        self.second_notice_var = tk.BooleanVar()
        self.second_notice_checkbox = ttk.Checkbutton(
            send_emails_frame,
            text="Second Notice",
            variable=self.second_notice_var,
            command=lambda: self.send_emails_btn.configure(
                state=tk.NORMAL if self.second_notice_var.get() else tk.DISABLED
            )
        )
        self.second_notice_checkbox.pack(side=tk.LEFT)

        self.go_live_var = tk.BooleanVar()
        self.go_live_checkbox = ttk.Checkbutton(
            send_emails_frame,
            text="GO LIVE",
            variable=self.go_live_var,
            command=lambda: self.send_emails_btn.configure(
                style='Red.TButton' if self.go_live_var.get() else 'TButton'
            )
        )
        self.go_live_checkbox.pack(side=tk.LEFT)

        # Create red button style
        style = ttk.Style()
        style.theme_use('alt')
        style.configure('Red.TButton', background='red')
        style.map('Red.TButton',
            background=[('active', '#ff0000'), ('!active', '#ff0000')],
            foreground=[('active', 'white'), ('!active', 'white')]
        )
    
    def on_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
    def on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_frame, width=event.width)

    def sort_jobs_missing_info_by_leader(self):
        leader_summary = self.df_jobs_missing_info.groupby('Leader').size().sort_values(ascending=False)
        self.df_jobs_missing_info['Leader'] = pd.Categorical(
            self.df_jobs_missing_info['Leader'],
            categories=leader_summary.index,
            ordered=True
        )
        self.df_jobs_missing_info = self.df_jobs_missing_info.sort_values('Leader')

    def find_missing_contacts(self):
        try:
            # Clear previous results
            self.df_jobs_missing_info = None
            self.missing_info_text.delete(1.0, tk.END)
            self.download_missing_info_btn.config(state=tk.DISABLED)
            self.send_emails_btn.config(state=tk.DISABLED)

            # Increase window height to accommodate missing info section
            self.root.geometry("700x500")

            contacts_files = {}
            for frame in self.contact_frames:
                company, file_path = frame.get_data()
                if not company or not file_path:
                    messagebox.showerror("Error", "Please complete all contact file entries")
                    return
                contacts_files[company] = file_path

            # Load all data
            contacts_dfs = load_job_contacts_files(contacts_files)
            df_invs = load_invoices(self.file_paths['invoices_file'].get())
            df_projects = load_projects(self.file_paths['projects_file'].get())
            df_emails = load_emails(self.file_paths['users_file'].get())

            is_lien_exports_invoice_file = 'order_no' in df_invs.columns
            is_hts_or_dxs_contacts_file = 'HTS' in contacts_files.keys() or 'DXS' in contacts_files.keys()
            
            if is_lien_exports_invoice_file and is_hts_or_dxs_contacts_file:
                response = messagebox.askyesno(
                    "Warning",
                    "The invoice file is pulled from Lien Exports (typically for ONCO) and there are HTS/DXS contacts files. Do you want to continue?"
                )
                if not response:
                    return

            df_job_contacts = create_job_contacts_file(contacts_dfs)
            df_job_contacts = filter_job_contacts_for_invoice_file(df_job_contacts, df_invs)
            df_job_contacts = append_leader_to_job_contacts(df_job_contacts, df_projects)
            df_job_contacts = append_emails_to_job_contacts(df_job_contacts, df_emails)
            self.df_job_contacts = df_job_contacts

            df_jobs_missing_info = filter_job_contacts_for_missing_info(df_job_contacts)

            existing_urls = pd.read_csv(PROJECT_LINKS_FILE)
            self.df_jobs_missing_info = attach_urls_to_job_contacts(df_jobs_missing_info, existing_urls)

            needs_new_urls = self.df_jobs_missing_info[self.df_jobs_missing_info['URL'].isna()]
            
            if needs_new_urls.shape[0] > 0:
                response = messagebox.askyesno(
                    "URL Scraping Required",
                    f"{needs_new_urls.shape[0]} projects need new URLs. Do you want to scrape them now?"
                )
                if response:
                    scrape_for_new_urls(needs_new_urls, existing_urls)
                    existing_urls = pd.read_csv(PROJECT_LINKS_FILE)

                self.df_jobs_missing_info = attach_urls_to_job_contacts(self.df_jobs_missing_info, existing_urls)

            # Create categorical Leader column with ordered categories from leader_summary
            self.sort_jobs_missing_info_by_leader()

            # Display Missing Info
            self.display_missing_info()

            # Enable buttons
            self.download_missing_info_btn.config(state=tk.NORMAL)
            self.send_emails_btn.config(state=tk.NORMAL)
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while processing files:\n{str(e)}")

    def upload_missing_info_file(self, file_path):
        missing_info_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv")]
        )        
        self.missing_info_path.set(missing_info_path)
        self.df_jobs_missing_info = pd.read_csv(missing_info_path)
        self.sort_jobs_missing_info_by_leader()

        self.df_job_contacts = None # Reset job contacts

        self.display_missing_info()
        self.send_emails_btn.config(state=tk.NORMAL)

    def display_missing_info(self):
        self.missing_info_text.delete(1.0, tk.END)
        leader_summary = self.df_jobs_missing_info.groupby('Leader').size().sort_values(ascending=False)

        if self.df_job_contacts is not None:
            self.missing_info_text.insert(tk.END, f"{self.df_jobs_missing_info.shape[0]} out of {self.df_job_contacts.shape[0]} projects missing information; See below by Leader:\n\n")
        else:
            self.missing_info_text.insert(tk.END, f"A Missing Info File Was Uploaded. {self.df_jobs_missing_info.shape[0]} projects missing information; See below by Leader:\n\n")

        for leader, count in leader_summary.items():
            self.missing_info_text.insert(tk.END, f"{leader}: {count} projects\n")

    def download_missing_info(self):
        try:
            # Open file dialog to get save location
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")]
            )
            
            if filename:
                # Save missing jobs to CSV
                self.df_jobs_missing_info.to_csv(filename, index=False)
                messagebox.showinfo("Success", "Mamaux contacts file saved successfully!")

                # Open the file
                os.startfile(filename)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving the file:\n{str(e)}")

    def send_emails(self):
        try:
            outlook = connect_to_outlook()
            if not outlook:
                messagebox.showerror("Error", "Could not connect to Outlook")
                return
            
            # Group ojects by leader
            leader_groups = self.df_jobs_missing_info.groupby('Leader')
            for leader, projects in leader_groups:
                if self.go_live_var.get():
                    to_email = projects['Leader Email'].iloc[0]
                else:
                    to_email = self._TEST_TO_EMAIL

                # Prepare email content
                subject = "PLEASE READ: Your Projects Requiring Contact Information Updates"
                second_notice_text = ' This is the second notification for the month. Reminder that lien notices are sent out on the 15th of the month (where applicable).' if self.second_notice_var.get() else ''
                body_text = f"""{leader.split(' ')[0]},

the following projects are missing required contact information.{second_notice_text} Please review and update where necessary (highlighted in yellow):"""
                print (leader)
                print (projects['Leader Email'].iloc[0])

                # Send email
                send_outlook_email(
                    outlook=outlook,
                    to_address=to_email,
                    cc_addresses='ryan.skene@hts.com',
                    subject=subject,
                    df=projects.drop(['Leader Email'], axis=1),
                    body_text_prescript=body_text,
                )
            
            self.go_live_var.set(False)
            messagebox.showinfo("Success", "All emails have been sent successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while sending emails:\n{str(e)}")



def main():
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()