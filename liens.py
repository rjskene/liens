import pandas as pd
import win32com.client

def load_job_contacts_files(
        contacts_files: dict[str, str]
):
    """
    Load the job contacts files
    """
    # One of the HTS files has a last row where the final column value is NA, when all other values in the column are 0
    # This generates a warning for unexpected end of data;
    contacts_dfs = {k: pd.read_csv(v, on_bad_lines='warn', encoding='utf-8', engine='python') for k, v in contacts_files.items()}

    return contacts_dfs

def load_invoices(file_path: str):
    return pd.read_csv(file_path)

def load_emails(file_path: str):
    return pd.read_csv(file_path)

def load_existing_jobs(file_path: str):
    df_existing_jobs = pd.read_excel(file_path, sheet_name='Data', header=1)
    df_existing_jobs['Job Number'] = df_existing_jobs['Job Number'].astype(str)
    return df_existing_jobs

def load_projects(file_path: str):
    df_projects = pd.read_csv(file_path)
    df_projects = df_projects.set_index('Project ID')
    return df_projects

def filter_invoice_file(
    df_invs: pd.DataFrame,
    df_job_contacts: pd.DataFrame
):
    """
    Filter the invoice file

    1. Exclude any projects where project ID begins with P
    2. Exclude any projects where project ID begins with I
    3. By excluding all "P" and "I" projects, all Job IDs should be 8 digits long
    """

    # jobs_id_bool = df_invs['Project ID'].isin(df_job_contacts['Project Number'])
    p_bool = ~(~df_invs['Project ID'].str.contains('-') & (df_invs['Project ID'].apply(lambda x: x[0] == 'P')))
    i_bool = ~(~df_invs['Project ID'].str.contains('-') & (df_invs['Project ID'].apply(lambda x: x[0] == 'I')))

    df_inv_hts_dxs = df_invs.loc[p_bool & i_bool].copy()
    df_inv_hts_dxs['Job ID'] = df_inv_hts_dxs['Project ID'].str.split('-').str[0]
    assert (df_inv_hts_dxs['Job ID'].str.len() == 8).all()

    missing_jobs = df_inv_hts_dxs.loc[~df_inv_hts_dxs['Project ID'].isin(df_job_contacts['Project Number'])]
    found_jobs = df_inv_hts_dxs.loc[df_inv_hts_dxs['Project ID'].isin(df_job_contacts['Project Number'])]

    return missing_jobs, found_jobs

def create_job_contacts_file(
    contacts_dfs: dict[str, pd.DataFrame],
):
    """
    Purpose: create a single, comprehensive job contacts file for all companies that includes Leader information
    1. Combine HTS & DXS Job Contact Files
    2. Add Job ID column
    3. remove any projects with a sub-decimal (e.g. 12345678.1); these seem to always result in duplicates, so drop them
    4. merge with projects file to add Leader column

    5. filter for jobs that have a Leader
    """
    for k, v in contacts_dfs.items():
        v.loc[:, 'Company'] = k

    df_job_contacts = pd.concat(contacts_dfs.values())

    df_job_contacts = df_job_contacts.reset_index(drop=True)
    df_job_contacts = df_job_contacts.fillna('')
    df_job_contacts.loc[:, 'Job ID'] = df_job_contacts['Project Number'].str.split('-').str[0]

    proj_id_has_sub_decimal = df_job_contacts['Project Number'].apply(lambda val: val[-2] == '.')
    df_job_contacts.loc[proj_id_has_sub_decimal, 'Project Number'] = df_job_contacts.loc[proj_id_has_sub_decimal, 'Project Number'].str.split('.').apply(lambda val: ''.join(val[:-1]))
    df_job_contacts =df_job_contacts.loc[df_job_contacts['Project Number'].drop_duplicates().index]

    assert df_job_contacts.set_index('Project Number').index.is_unique

    return df_job_contacts

def filter_job_contacts_for_invoice_file(
    df_job_contacts: pd.DataFrame,
    df_invs: pd.DataFrame
):
    """
    Filter the job contacts for the invoice file
    """

    df_job_contacts = df_job_contacts.loc[df_job_contacts['Project Number'].isin(df_invs['Project ID'])].copy()

    return df_job_contacts

def convert_job_list_to_mamaux_format(
    df_job_list: pd.DataFrame
):
    """
    Convert the job list to the format that Jeff's team uses
    """

    col_names = {
        'Project Nickname': 'Job Name',
        'Owner Name': 'Owner Name',
        'Owner Address': 'Owner Address', 
        'GC Name': 'General Contractor (GC) Name',
        'GC Address': 'GC Address',
        'Customer Name': 'Mechanical Contractor (MC) Name',
        'Customer Address': 'MC Address',
        'Customer City': 'MC City',
        'Customer State': 'MC State',
        'Customer Zip': 'MC Zip',
        'Job ID': 'Job Number',
    }
    df_job_list = df_job_list.rename(columns=col_names)
    df_job_list['Job ID'] = df_job_list['Project Number'].str.split('-').str[0]
    cols = [
        'Job Number',
        'Company',
        'Job Name', 
        'Owner Name',
        'Owner Address',
        'Owner City',
        'Owner State', 
        'Owner Zip',
        'General Contractor (GC) Name',
        'GC Address',
        'GC City',
        'GC State',
        'GC Zip',
        'Mechanical Contractor (MC) Name',
        'MC Address',
        'MC City',
        'MC State',
        'MC Zip',
    ]
    return df_job_list[cols]

def filter_mamaux_contacts_for_existing_jobs(
    df_mamaux_contacts: pd.DataFrame,
    df_existing_jobs: pd.DataFrame
):
    """
    Filter the job contacts for the existing jobs
    """
    df_mamaux_contacts = df_mamaux_contacts.loc[~df_mamaux_contacts['Job Number'].isin(df_existing_jobs['Job Number'])].copy()

    return df_mamaux_contacts

def append_missing_jobs_to_mamaux_contacts(
    df_mamaux_contacts: pd.DataFrame,
    df_job_contacts: pd.DataFrame,
    df_invs: pd.DataFrame
):
    df_filtered_invs = df_invs.loc[~(
        df_invs['Project ID'].str.contains('VRFS') 
        | df_invs['Project ID'].str.contains('ONCO')
        | df_invs['Project ID'].apply(lambda x: x.startswith('P'))
        | df_invs['Project ID'].apply(lambda x: x.startswith('I'))
    )]

    missing_invs = df_filtered_invs.loc[~df_filtered_invs['Project ID'].isin(df_job_contacts['Project Number'])].copy()

    missing_invs['Job Number'] = missing_invs['Project ID'].str.split('-').str[0]
    missing_invs = missing_invs.rename(columns={'Job Name': 'Project Title'})

    # Add missing jobs to mamaux contacts
    for _, row in missing_invs.iterrows():
        new_row = pd.DataFrame({
            'Job Number': [row['Job Number']],
            'Company': ['MANUAL ENTRY REQUIRED'],  # Assuming HTS as default
            'Job Name': [row['Project Title']],
            # Fill remaining columns with empty values
            'Owner Name': [''],
            'Owner Address': [''],
            'Owner City': [''],
            'Owner State': [''], 
            'Owner Zip': [''],
            'General Contractor (GC) Name': [''],
            'GC Address': [''],
            'GC City': [''],
            'GC State': [''],
            'GC Zip': [''],
            'Mechanical Contractor (MC) Name': [''],
            'MC Address': [''],
            'MC City': [''],
            'MC State': [''],
            'MC Zip': ['']
        })
        df_mamaux_contacts = pd.concat([df_mamaux_contacts, new_row], ignore_index=True)

    return df_mamaux_contacts

def append_leader_to_job_contacts(
    df_job_contacts: pd.DataFrame,
    df_projects: pd.DataFrame
):
    """
    Purpose: append the Leader column to the job contacts file
    """
    df_job_contacts = pd.merge(
        df_job_contacts.set_index('Project Number'),
        df_projects['Leader'],
        how='left',
        left_index=True,
        right_index=True,
    ) # incorporate Leader column into jobs file
    df_job_contacts = df_job_contacts.reset_index(drop=False)

    df_job_contacts = df_job_contacts.loc[~df_job_contacts.Leader.isna()]

    return df_job_contacts

def append_emails_to_job_contacts(
    df_job_contacts: pd.DataFrame,
    df_emails: pd.DataFrame
):
    """
    Append the emails to the job contacts file
    """

    df_emails.loc[:, 'Name'] = df_emails[['First Name', 'Surname']].apply(lambda x: ' '.join(x), axis=1)
    assert df_job_contacts['Leader'].isin(df_emails['Name']).all(), f"Leader {df_job_contacts['Leader'].isin(df_emails['Name']).sum()} not found in emails"
    df_job_contacts.loc[:, 'Leader Email'] = df_emails.set_index('Name').loc[df_job_contacts['Leader']]['Email'].values

    return df_job_contacts


def filter_job_contacts_for_missing_info(
        df_job_contacts: pd.DataFrame,
):
    """
    Generate the job list that is used by Jeff team to create lien letter

    ### NOTE ###
    + the Mechanical Contractor columns are set assuming that the customer is the MC; this is the convention Jeff seems to use

    1. Create a copy of the job contacts file
    2. Filter out specific columns
    3. Filter for jobs that are missing information
    """

    df_job_list = df_job_contacts.copy()

    cols = [
        'Company',
        'Project Number',
        'Project Nickname',
        'Customer Name', 'Customer Phone',
        'Customer Address', 'Customer City', 'Customer State', 'Customer Zip',
        'Customer Role', 'GC Name', 'GC Phone', 'GC Address', 'GC City',
        'GC State', 'GC Zip', 'Owner Name', 'Owner Phone',
        'Owner Address', 'Owner City', 'Owner State',
        'Owner Zip', 'Leader', 'Leader Email'
    ]
    df_job_list = df_job_list[cols]
    contact_cols = df_job_list.columns.str.contains('Owner') | df_job_list.columns.str.contains('GC')
    
    df_job_list = df_job_list.where(df_job_list != '', None)
    df_job_list = df_job_list.where(df_job_list != ' ', None)
    is_missing_values = (df_job_list.loc[:, contact_cols] == '') | (df_job_list.loc[:, contact_cols] == ' ') | df_job_list.loc[:, contact_cols].isna()
    df_jobs_missing_info = df_job_list.loc[is_missing_values.loc[is_missing_values.any(axis=1)].index].copy()

    return df_jobs_missing_info

def connect_to_outlook():
    """Establish connection to Outlook"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook
    except Exception as e:
        print(f"Error connecting to Outlook: {e}")
        return None

def df_to_html_table(df: pd.DataFrame) -> str:
    """Convert DataFrame to HTML table with basic styling"""
    df = df.where(~df.isna(), 'Missing!')

    def style_missing(val):
        if val == 'Missing!':
            return 'background-color: yellow; color: red'
        return ''
    
    styled_df = df.style.map(style_missing)
    df = styled_df
    styles = """
        <style>
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 1px solid black; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
        </style>
    """
    html_table = df.to_html(index=False)
    return f"{styles}\n{html_table}"

def send_outlook_email(
    outlook: win32com.client.Dispatch,
    to_address: str,
    subject: str,
    df: pd.DataFrame,
    body_text: str = "",
    cc_addresses: list[str] = [],
) -> None:
    """
    Send an email through Outlook with a DataFrame displayed as an HTML table
    
    Args:
        to_addresses: List of email addresses to send to
        subject: Email subject line
        df: DataFrame to include in email body
        body_text: Optional text to include before the table
        cc_addresses: Optional list of email addresses to CC
    """
    # Create Outlook application object
    cc_addresses += ['ryan.skene@hts.com']

    mail = outlook.CreateItem(0)  # 0 represents olMailItem
    
    # Set email properties
    mail.Subject = subject
    mail.To = to_address
    # if cc_addresses:
    #     mail.CC = "; ".join(cc_addresses)
    
    # Create HTML body with optional text and table
    html_table = df_to_html_table(df)
    mail.HTMLBody = f"{body_text}<br><br>{html_table}"
    
    # Send the email
    mail.Send()


