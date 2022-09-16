from __future__ import print_function
#to create the different templates
from docxtpl import DocxTemplate
import sys

#Time Delay
import time



#Create df's from csv's
import pandas as pd
import os

#Identify file directories of output pdf files
import os.path

#Setup gmail,secure connection
import smtplib,ssl

#Maskpassword
import maskpass

#Progress bar
import tqdm
import tqdm

from tqdm import tqdm

#retrieves all PDF files from app directory 
from collections import ChainMap

#building Email Parts
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText


#Setting up gmail connections (between app and gmail)
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(os.path.realpath(sys.executable))
elif __file__:
    application_path = os.path.dirname(__file__)

# determine if application is a script file or frozen exe


def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        results = service.users().labels().list(userId='me').execute()
        labels = results.get('labels', [])

        

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')


if __name__ == '__main__':
    main()

#Initialization of MS Word
from win32com import client
word_app = client.Dispatch("word.application")

print("Mail Merging Initializing")

Data_frame = pd.read_csv("commissions_output.csv") # whole source document
Data_frame1=pd.read_csv("commissions_output.csv", usecols= ['email']) # where the emails are found
Data_frame2=pd.read_csv("commissions_output.csv", usecols= ['file']) #shoo pdf files are named

#translates Data Frames 1 & 2 to List
mailtos=Data_frame1.values.tolist() # contains email addresses of recepients
pdf_files=Data_frame2.values.tolist() # contains pdf file names

# Identifying file directory and adding names as keys to dictionary
my_files =[{each_file.split(".")[0]:each_file} for each_file in os.listdir(os.path.abspath("PDF/")) if each_file.endswith(".pdf")]
my_files_dict = dict(ChainMap(*my_files))

#------------STEP 1-Append-records from source(.csv) to doc template and convert to pdf to be saved in "pdf/" folder

for r_index, row in tqdm(Data_frame.iterrows(), total=Data_frame.shape[0]):
    agent_name =row['Name'] # for file Name
    tpl = DocxTemplate("mailmergev2.docx") #in the same directory
    df_to_doct = Data_frame.to_dict() # dataframe ->dict for the template render
    x = Data_frame.to_dict(orient='records')
    context = x
    tpl.render(context[r_index]) 
    tpl.save('Doc\\'+agent_name+".docx")
    

    #get project folder path # os.path.dirname(os.path.abspath(__file__))
    root_dir=application_path
  
    

    # Convert file from Docx to PDF
    doc = word_app.documents.Open(root_dir+ '\\Doc\\' + agent_name + '.docx')
  
    doc.SaveAs(root_dir+'\\PDF\\'+ agent_name + '.pdf', FileFormat=17)
    doc.Close()
print("Mail Merging Completed. Generating Emails...")
word_app.Quit()


print("Input Month and Year in the next line ")


#----------------STEP-2-Establish-secure-connection-with-gmail-(SSL)
port = 465  # For SSL
sender_email = "owinobasildev@gmail.com"
Subject = input() 
password = maskpass.askpass(prompt="Enter Password",mask="*")


# Create a secure SSL context using localhost's CA and security certifications
context = ssl.create_default_context()

for i in tqdm(range(len(mailtos))):
  

    name = pdf_files [i]
    panga= mailtos [i] [0]

    filename = pdf_files [i] [0]
    msg2 = MIMEMultipart()
    msg2['Subject'] = "AFS Sales Summary | " + Subject 
    msg2['From'] = sender_email
    msg2['To'] = panga 
    
    # HTML Version of the message
    
    html = """\
    <html>
    <body>
        <p>Greetings,<br>
        Please find attached the above-mentioned for you reference<br>
        
        Best Regards
        </p>
    <table style="width: 444px; font-size: 9pt; font-family: Arial, sans-serif; line-height:normal;" cellpadding="0" cellspacing="0">
    <tbody>
        <tr>
            <td colspan="2" style=" font-size:12pt; font-family:Arial,sans-serif; padding-bottom:6px;  width:444px">
                <span style="font-family: Arial, sans-serif; color:#1b7592; font-weight: bold">
                    
                    <span style="font-family: Arial, sans-serif; color:#1b7592; font-weight: bold"> | </span>
                </span>
                <span style="font-family: Arial, sans-serif; color:#1b7592; font-weight: bold">Payroll</span>
            </td>
        </tr>
        <tr>
            <td style="width: 289px; padding-top: 6px; padding-bottom:6px;">
                <table cellpadding="0" cellspacing="0">
                    <tbody>
                    
                    <tr>
                        <td><span style="font-size: 9pt; line-height: 13pt; font-family: Arial, sans-serif; color:#100000;">Email:</span></td>
                        <td><span><a href="mailto:hr@avanzar.ke" style="font-size: 9pt; line-height: 13pt; font-family: Arial, sans-serif; color:#100000; font-weight: bold; text-decoration: none"><span style="font-size: 9pt; line-height: 13pt; font-family: Arial, sans-serif; color:#100000; font-weight: bold; text-decoration: none">hr@avanzar.ke</span></a></span></td>
                    </tr>
                    
                </tbody></table>
            </td>
            
            <td style="width:155px; padding-top: 6px; padding-bottom:6px; text-align:right;">
                
                <img border="0" alt="Logo" height="56" style="max-width: 100%; width:auto; height:36px; border:0;"  src="https://drive.google.com/uc?export=view&amp;id=1M6RevA_TzgWCOtIRXlz4gWcEMcOLdnri">
            </td>
            
        </tr>
        <tr>
            <td style="border-top: 1px solid #1b7592" colspan="2" width="444">
        </td></tr>
        <tr>
            <td style="padding-top: 10px; padding-bottom: 10px; font-size: 10pt; font-family: Arial, sans-serif; font-weight:bold; color: #1793ce; width:299px;" width="299">
                
            </td>
            <td style="width:145px; text-align:right; padding-top: 10px; padding-bottom: 10px; " width="145">
                <span style="display:inline-block; height:19px;"><span><a href="https://www.facebook.com/MyCompany" target="_blank"><img alt="Facebook icon" border="0" width="19" height="19" style="border:0; height:19px; width:19px" src="https://www.mail-signatures.com/signature-generator/img/templates/top-security/fb.png"></a>&nbsp;&nbsp;</span><span><a href="https://www.linkedin.com/company/mycompany404" target="_blank"><img alt="LinkedIn icon" border="0" width="19" height="19" style="border:0; height:19px; width:19px" src="https://www.mail-signatures.com/signature-generator/img/templates/top-security/ln.png"></a>&nbsp;&nbsp;</span><span><a href="https://twitter.com/MyCompany404" target="_blank"><img alt="Twitter icon" border="0" width="19" height="19" style="border:0; height:19px; width:19px" src="https://www.mail-signatures.com/signature-generator/img/templates/top-security/tt.png"></a>&nbsp;&nbsp;</span><span><a href="https://www.youtube.com/user/MyCompanyChannel" target="_blank"><img alt="Youtube icon" border="0" width="19" height="19" style="border:0; height:19px; width:19px" src="https://www.mail-signatures.com/signature-generator/img/templates/top-security/yt.png"></a>&nbsp;&nbsp;</span><span><a href="https://www.instagram.com/mycompany404/" target="_blank"><img alt="Instagram icon" border="0" width="19" height="19" style="border:0; height:19px; width:19px" src="https://www.mail-signatures.com/signature-generator/img/templates/top-security/it.png"></a>&nbsp;&nbsp;</span><span><a href="https://pinterest.com/mycompany404/" target="_blank"><img alt="Pinterest icon" border="0" width="19" height="19" style="border:0; height:19px; width:19px" src="https://www.mail-signatures.com/signature-generator/img/templates/top-security/pt.png"></a></span></span>
            </td>
        </tr>
        
        <tr>
            <td style="font-size:8pt; line-height: 9pt; font-family:Arial,sans-serif; color:#878787; width:444px; text-align:justify;" colspan="2" width="444">
                <span style="font-family: Arial, sans-serif; color:#000000">The content of this email is confidential and intended for the recipient specified in message only. It is strictly forbidden to share any part of this message with any third party, without a written consent of the sender. If you received this message by mistake, please reply to this message and follow with its deletion, so that we can ensure such a mistake does not occur in the future.</span>
            </td>
        </tr>
        </tbody>
        </table>    
    </body>
    </html>
    """

    # Turn these into plain/html MIMEText objects
    part2 = MIMEText(html, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    msg2.attach(part2)
        
    fo = open(root_dir + '\\PDF\\' + filename ,'rb') 
    attachfile = MIMEApplication(fo.read(),_subtype="pdf")
    fo.close()
    attachfile.add_header('Content-Disposition','attachment',filename=filename)
    msg2.attach(attachfile)
    s = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context)
    s.login("owinobasildev@gmail.com", password)    
    s.sendmail(sender_email,[panga],msg2.as_string())  
            
print('Emails Sent Successfully')
s.close()    
time.sleep(30)