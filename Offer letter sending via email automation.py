import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO
import smtplib
import imaplib
import email
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime

excel_file_path = r"Sample.xlsx"    #path to the file
df = pd.read_excel(excel_file_path)


hrm_pdf_path = r"HRM INTERNSHIP OFFER LETTER.pdf"    #path to the file
marketing_pdf_path = r"MAREKTING INTERNSHIP OFFER LETTER.pdf"    #path to the file


from_email = 'Internship@holograd.in'    #Hostinger email
password = ''    #email password

#PDF logic
def create_custom_pdf(name, email_id, pdf_template_path, output_pdf_path):
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont("Helvetica-Bold", 12)
    can.drawString(55, 644, name)     #coordinates to place name
    can.drawString(55, 626, email_id)     #coordinates to place email
    current_date = datetime.now().strftime("%d %B %Y")
    can.drawString(426, 644, current_date)     #coordinates to place date
    can.save()

    packet.seek(0)
    overlay_pdf = PdfReader(packet)
    
    pdf_reader = PdfReader(pdf_template_path)
    pdf_writer = PdfWriter()


    for page_number in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_number]


        if page_number == 0:
            page.merge_page(overlay_pdf.pages[0])

        pdf_writer.add_page(page)


    with open(output_pdf_path, 'wb') as output_pdf_file:
        pdf_writer.write(output_pdf_file)

#Email logic
def send_email(subject, body, to_email, attachment_path):
    filename = os.path.basename(attachment_path)

    #email compose
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    #PDF attachment
    with open(attachment_path, 'rb') as attachment_file:
        part = MIMEApplication(attachment_file.read(), Name=filename)
        part['Content-Disposition'] = f'attachment; filename="{filename}"'
        msg.attach(part)

    #SMTP mail sending
    with smtplib.SMTP('smtp.hostinger.com', 587) as server:  # Ensure SMTP settings are correct
        server.starttls()
        server.login(from_email, password)
        server.sendmail(from_email, to_email, msg.as_string())

    #IMAP sent mail saved 
    save_sent_email(from_email, password, msg)

#IMAP sent mail saved to sent inbox 
def save_sent_email(from_email, password, msg):
    with imaplib.IMAP4_SSL('imap.hostinger.com') as imap:  # Ensure IMAP settings are correct
        imap.login(from_email, password)
        
        imap.select('"INBOX.Sent"')  # Adjust the folder name if necessary
        
        imap.append('INBOX.Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), str(msg).encode('utf-8'))
        
        imap.logout()

email_count = 0

#excel file processing
for index, row in df.iterrows():
    name = row['Name']
    email_id = row['Email ID']
    domain = row['Domain']
    
    if domain.lower() == 'hr':
        pdf_path = hrm_pdf_path
        output_pdf_path = rf"HRM INTERNSHIP OFFER LETTER.pdf"    #path to the file
        subject = 'Selection for HRM Internship'
        body = '''

--

Dear Intern,

I hope this email finds you well and filled with anticipation for the exciting journey ahead.

I am delighted to extend to you an offer to join us as an intern at HoloGrad. Congratulations on your selection! Your exceptional skills and passion for innovation have impressed us, and we are eager to welcome you to our team.

Please find attached your official offer letter, outlining the terms and conditions of your internship with us.

Additionally, I would like to inform you about the training task that we have prepared for you.
This task will provide you with valuable insights into the work you will be undertaking during your internship and will help you familiarize yourself with our projects and methodologies.

Your team leader will be shortly contacting you to discuss the details of the training task and to provide you with any assistance or guidance you may require. We encourage you to approach this task with enthusiasm and curiosity, as it will serve as an excellent opportunity for you to showcase your abilities and learn from the experience.

Once again, congratulations on your selection for the internship at HoloGrad. We are thrilled to have you on board and look forward to working closely with you.

If you have any questions or require further clarification, please do not hesitate to reach out to us.

Warm regards,

HR Department
HoloGrad
'''

    elif domain.lower() == 'marketing':
        pdf_path = marketing_pdf_path
        output_pdf_path = rf"MARKETING INTERNSHIP OFFER LETTER.pdf"    #path to the file
        subject = 'Selection for Marketing Internship'
        body = '''

--

Dear Intern,

I hope this email finds you well and filled with anticipation for the exciting journey ahead.

I am delighted to extend to you an offer to join us as an intern at HoloGrad. Congratulations on your selection! Your exceptional skills and passion for innovation have impressed us, and we are eager to welcome you to our team.

Please find attached your official offer letter, outlining the terms and conditions of your internship with us.

Additionally, I would like to inform you about the training task that we have prepared for you.
This task will provide you with valuable insights into the work you will be undertaking during your internship and will help you familiarize yourself with our projects and methodologies.

Your team leader will be shortly contacting you to discuss the details of the training task and to provide you with any assistance or guidance you may require. We encourage you to approach this task with enthusiasm and curiosity, as it will serve as an excellent opportunity for you to showcase your abilities and learn from the experience.

Once again, congratulations on your selection for the internship at HoloGrad. We are thrilled to have you on board and look forward to working closely with you.

If you have any questions or require further clarification, please do not hesitate to reach out to us.

Warm regards,

HR Department
HoloGrad
'''
    else:
        print(f"Unknown domain for {name}, skipping...")
        continue

    create_custom_pdf(name, email_id, pdf_path, output_pdf_path)

    send_email(subject, body, email_id, output_pdf_path)

    print(f"PDF for {name} ({domain}) has been updated and sent to {email_id}.")

    email_count += 1

print("Processing complete.")
print("Total emails sent:",email_count)
