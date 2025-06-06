import os
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

    Must distinguish between the two different invoice files
        + AR file has a column 'Project ID'
        + Lien exports file has a column 'order_no
    """
    if 'Project ID' in df_invs.columns:
        df_job_contacts = df_job_contacts.loc[df_job_contacts['Project Number'].isin(df_invs['Project ID'])].copy()
    elif 'order_no' in df_invs.columns:
        df_job_contacts = df_job_contacts.loc[df_job_contacts['Project Number'].isin(df_invs['order_no'])].copy()
    else:
        raise ValueError('Project Number or order_no column not found in job contacts')

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
        'Customer Role', 'GC Name', 'GC Address', 'GC City',
        'GC State', 'GC Zip', 'Owner Name',
        'Owner Address', 'Owner City', 'Owner State',
        'Owner Zip', 'Leader', 'Leader Email'
    ]
    
    df_job_list = df_job_list[cols]
    contact_cols = df_job_list.columns.str.contains('Owner') | df_job_list.columns.str.contains('GC')
    
    df_job_list = df_job_list.where(df_job_list != '', None)
    df_job_list = df_job_list.where(df_job_list != ' ', None)
    is_missing_values = (df_job_list.loc[:, contact_cols] == '') | (df_job_list.loc[:, contact_cols] == ' ') | df_job_list.loc[:, contact_cols].isna()
    df_jobs_missing_info = df_job_list.loc[is_missing_values.loc[is_missing_values.any(axis=1)].index].copy()

    df_jobs_missing_info = df_jobs_missing_info.loc[~(df_jobs_missing_info['Customer Name'] == df_jobs_missing_info['Owner Name'])]
    df_gc_is_customer = df_jobs_missing_info.loc[(df_jobs_missing_info['Customer Name'] == df_jobs_missing_info['GC Name'])].copy()
    contact_cols = df_jobs_missing_info.columns.str.contains('Owner')
    is_df_gc_is_customer_missing_values = (df_gc_is_customer.loc[:, contact_cols] == '') | (df_gc_is_customer.loc[:, contact_cols] == ' ') | df_gc_is_customer.loc[:, contact_cols].isna()
    df_jobs_missing_info = df_jobs_missing_info.loc[df_jobs_missing_info.index.difference(df_gc_is_customer.index)]
    df_jobs_missing_info = pd.concat([df_jobs_missing_info, df_gc_is_customer.loc[is_df_gc_is_customer_missing_values.any(axis=1)]])

    assert ~(df_jobs_missing_info['Customer Name'] == df_jobs_missing_info['Owner Name']).any()
    assert df_jobs_missing_info.loc[df_jobs_missing_info['Customer Name'] == df_jobs_missing_info['GC Name']].loc[:, df_jobs_missing_info.columns.str.contains('Owner')].isna().any(axis=1).all()
    
    return df_jobs_missing_info

def attach_urls_to_job_contacts(
    df_jobs_missing_info: pd.DataFrame,
    existing_urls: pd.DataFrame
):
    """
    Attach the URLs to the job contacts file
    """
    if not all(col in existing_urls.columns for col in ["Project Number", "URL"]):
        raise ValueError("existing_urls DataFrame must contain only 'Project Number' and 'URL' columns")
    
    df_copy = df_jobs_missing_info.copy()
    if 'URL' in df_copy.columns:
        df_copy = df_copy.drop(columns=['URL'])

    df_jobs_missing_info = pd.merge(df_copy, existing_urls, on='Project Number', how='left')
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
    # Create a copy to avoid modifying the original dataframe

    if df.empty:
        return '<p>No data to display</p>'

    df_copy = df.copy().reset_index(drop=True)
    df_copy = df_copy.drop(columns=['Company'])
    
    # Check if URL column exists
    has_url = 'URL' in df_copy.columns
    if not has_url:
        raise ValueError('URL column not found')
        
    # Mark missing values
    df_copy = df_copy.where(~df_copy.isna(), 'Missing!')

    def style_missing(val):
        if val == 'Missing!':
            return 'background-color: yellow; color: red'
        return ''
    
    styled_df = df_copy.style.map(style_missing)

    df_copy = df_copy.set_index('Project Number')
    
    # Add hyperlinks to Project Number if URL column exists
    def make_clickable(project_id):
        if project_id in df_copy.index:
            return f'<a href="{df_copy.loc[project_id, "URL"]}">{project_id}</a>'
        return project_id
    
    styled_df = styled_df.format({'Project Number': make_clickable})
    
    styles = """
        <style>
            table { border-collapse: collapse; width: 100%; table-layout: fixed; }
            th, td { border: 1px solid black; padding: 10px; text-align: left; width: 400px; }
            th { background-color: #f2f2f2; }
            a:hover { cursor: pointer; }
        </style>
    """

    html_table = styled_df.to_html(index=False, escape=False)
    return f"{styles}\n{html_table}"

def send_outlook_email(
    outlook: win32com.client.Dispatch,
    to_address: str,
    subject: str,
    df: pd.DataFrame,
    body_text_prescript: str = "",
    body_text_postscript: str = "",
    cc_addresses: list[str] = None,
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
    if cc_addresses is None:
        cc_addresses = []
    elif isinstance(cc_addresses, str):
        cc_addresses = [cc_addresses]
        
    # Make a copy to avoid modifying the input parameter
    cc_list = cc_addresses.copy()

    mail = outlook.CreateItem(0)  # 0 represents olMailItem
    
    # Set email properties
    mail.Subject = subject
    mail.To = to_address
    if cc_list:
        mail.CC = "; ".join(cc_list)
    
    # Add the GIF as an embedded image in the HTML body
    must_be_completed_path = os.path.join(os.path.dirname(__file__), 'static', 'Must Be Completed.png')
    gif_path = os.path.join(os.path.dirname(__file__), 'static', 'ProjectContactsVid17s.gif')
    customer_is_gc_path = os.path.join(os.path.dirname(__file__), 'static', 'If GC is Customer.png')
    customer_is_owner_path = os.path.join(os.path.dirname(__file__), 'static', 'If Owner is Customer.png')
    selections_correct_path = os.path.join(os.path.dirname(__file__), 'static', 'Selections Correct BUT.png')
    company_profile_incomplete_path = os.path.join(os.path.dirname(__file__), 'static', 'Company Profile Incomplete.png')

    # Add the image as an attachment with a Content ID
    attachment = mail.Attachments.Add(must_be_completed_path)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MustBeCompleted")
 
    attachment = mail.Attachments.Add(gif_path)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "myGIF")

    attachment = mail.Attachments.Add(customer_is_gc_path)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "customer_is_gc")

    attachment = mail.Attachments.Add(customer_is_owner_path)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "customer_is_owner")

    attachment = mail.Attachments.Add(selections_correct_path)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "selections_correct")

    attachment = mail.Attachments.Add(company_profile_incomplete_path)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "company_profile_incomplete")

    # Update the HTML body to include the embedded image
    html_table = df_to_html_table(df)
    mail.HTMLBody = f"""
        {body_text_prescript}<br><br>
        {html_table}<br><br>
        Please contact me if you have any comments / questions.<br><br>
        <b><u>REMINDERS & TIPS:</u></b><br><br>
        1. The top rows on the Contacts page must be completed for all projects.<br><br>
        <img src='cid:MustBeCompleted'><br><br>
        In most cases, the steps are as simple as below:<br><br>
        <img src='cid:myGIF'><br><br>
        2. Only complete Bonder if the job is bonded. Reminder that the Bonder is the bonding surety (not the customer/purchaser/owner/designer/etc.)<br><br>
        3.	If your Customer is the GC, please still select the customer as the Main GC in the top slot.<br><br>
        <img src='cid:customer_is_gc'><br><br>
        4.	For Direct-to-Owner sales, please select the customer in both the Main Owner and Main GC slots.<br><br>
        <img src='cid:customer_is_owner'><br><br>
        On a rare occasion, the top dialog slot has been completed correctly, but the company profile is missing information, such as below:<br><br>
        <img src='cid:selections_correct'><br><br>
        <img src='cid:company_profile_incomplete'><br><br>
        {body_text_postscript}
    """

    # Send the email
    mail.Send()
