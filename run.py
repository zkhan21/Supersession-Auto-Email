import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

file_path = r'C:\Users\khz2okb\Desktop\Automation\CW44_Supersession_Tracker.xlsx'
email_file_path = r'C:\Users\khz2okb\Desktop\Automation\emails2.xlsx'

df = pd.read_excel(file_path, sheet_name='Sheet1')
email_df = pd.read_excel(email_file_path)

df[['FirstName', 'LastName']] = df['Responsible Name'].str.extract(r'(\w+) (\w+)')

merged_df = pd.merge(df, email_df, on=['FirstName', 'LastName'], how='left')

merged_df = merged_df[(merged_df['Predecessor Stock Value'].notnull()) & (merged_df['Predecessor Stock Value'] != 0)]

grouped = merged_df.groupby(['EMail'])

columns_to_extract = [
    "PredecessorBoschPn", "Sum of PredStockQty", "Predecessor Stock Value",
    "Customer", "ProjectID", "Description", "PredStockDisp",
    "SAPReasonCodeLongText", "Predecessor Escalation Level", "Sum of DaysRemainingOnPred"
]

smtp_server = 'rb-smtp-int.bosch.com'
smtp_port = 25
sender_email = 'Zuhair.Khan@us.bosch.com'
sender_password = 'Sanrrt088888'  # Change this to your actual password

for email, group_data in grouped:
    subject = 'Weekly Supersession Tracker'
    
    table_html = group_data[columns_to_extract].to_html(index=False)
    
    signature = """
    <p>Best Regards,<br><br>
    <strong>Zuhair Khan</strong><br>
    PAE Co-Op, Automotive Aftermarket, (AA-TG/RBR21-NA)<br>
    Robert Bosch LLC | Oakbrook Terrace Tower | 1 Tower Lane | Suite 3100 | Oakbrook Terrace, IL 60181 | USA | <a href="http://www.bosch.us">www.bosch.us</a><br>
    <a href="mailto:Zuhair.Khan@us.bosch.com">Zuhair.Khan@us.bosch.com</a></p>
    """
    
    what_to_do = """
    <p><strong>What to do:</strong></p>
    <ul>
        <li>Please ensure the predecessor stock is depleted before you reach the Escalation to EC.</li>
        <li>Please make sure you enter your notes such as current status, next steps, and target completion date in the Disposition Notes in the Project tab in Leap.</li>
    </ul>
    """
    
    links = """
    <p><strong>Links:</strong></p>
    <ul>
        <li><a href="https://leapna.bosch.com/">Link to Leap</a></li>
        <li><a href="https://leapna.bosch.com/leap/">Link to Task Queue</a></li>
        <li><a href="http://nais-bi.us.bosch.com/Reports/powerbi/LEAP/Supersession%20Tracker%20Leap">Weekly Supersession Tracker in Leap</a></li>
    </ul>
    <p><font size="2" color="lightgrey"><i>Note: This is an automated weekly email</i></font></p>
    """
    
    body = f"""\
    <html>
        <body>
            <p>Hello,</p>
            <p>Here are the supersession projects that require feedback for open dispositions:</p>
            {table_html}
            <br>
            {what_to_do}
            <br>
            {links}
            {signature}
        </body>
    </html>
    """
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))

    if not pd.isnull(email):
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.sendmail(sender_email, email, msg.as_string())
