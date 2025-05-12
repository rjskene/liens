import os
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
)
from mamaux_contacts_app import FileUploadBase, FileUploadMixin

class EmailApp(FileUploadMixin, FileUploadBase):
    def __init__(self, root):
        self.root = root
        self.root.title("Project Contact Emails")
        self.root.geometry("700x250")
        
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

        # Process button
        ttk.Button(
            upload_frame, 
            text="Find Missing Contacts", 
            command=self.find_missing_contacts,
            width=30,
        ).grid(row=6, column=0, columnspan=3, pady=20)

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
        # button_frame.grid_columnconfigure(0, weight=1)

        # Process button
        self.download_missing_info_btn = ttk.Button(
            button_frame, 
            text="Download Missing Info", 
            command=self.download_missing_info,
            state=tk.DISABLED,
            width=30,
        )
        self.download_missing_info_btn.grid(row=0, column=1, padx=50, pady=20)

        # Send emails button
        self.send_emails_btn = ttk.Button(
            button_frame, 
            text="Send Emails", 
            command=self.send_emails, 
            state=tk.DISABLED
        )
        self.send_emails_btn.grid(row=1, column=1, padx=50, pady=20)
        
    def on_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
    def on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_frame, width=event.width)

    def find_missing_contacts(self):
        try:
            # Clear previous results
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

            df_job_contacts = create_job_contacts_file(contacts_dfs)
            df_job_contacts = filter_job_contacts_for_invoice_file(df_job_contacts, df_invs)
            df_job_contacts = append_leader_to_job_contacts(df_job_contacts, df_projects)
            df_job_contacts = append_emails_to_job_contacts(df_job_contacts, df_emails)

            self.df_jobs_missing_info = filter_job_contacts_for_missing_info(df_job_contacts)

            # Display Missing Info
            leader_summary = self.df_jobs_missing_info.groupby('Leader').size().sort_values(ascending=False)
            self.missing_info_text.insert(tk.END, "Number of projects missing information:\n\n")
            for leader, count in leader_summary.items():
                self.missing_info_text.insert(tk.END, f"{leader}: {count} projects\n")

            # Enable buttons
            self.download_missing_info_btn.config(state=tk.NORMAL)
            self.send_emails_btn.config(state=tk.NORMAL)
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while processing files:\n{str(e)}")

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
            
            # Group projects by leader
            leader_groups = self.df_jobs_missing_info.groupby('Leader')
            
            for leader, projects in leader_groups:
                # email = projects['Leader Email'].iloc[0]
                email = 'ryan.skene@hts.com'
                
                # Prepare email content
                subject = "Projects Requiring Lien Information Updates"
                body_text = f"""Hello {leader},

The following projects are missing required lien information. Please review and update the missing information (highlighted in yellow) for your projects:"""
                
                # Send email
                send_outlook_email(
                    outlook=outlook,
                    to_address='ryan.skene@hts.com',
                    subject=subject,
                    df=projects.drop(['Leader Email'], axis=1),
                    body_text=body_text,
                )
            
            messagebox.showinfo("Success", "All emails have been sent successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while sending emails:\n{str(e)}")



def main():
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()