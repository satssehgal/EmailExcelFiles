import pandas as pd
import openpyxl
import os
import base64
from sendgrid.helpers.mail import (
    Mail, Attachment, FileContent, FileName,
    FileType, Disposition, ContentId)
from sendgrid import SendGridAPIClient
import pathlib

def send_email(dept, to_address, file):
	message = Mail(
	from_email=os.environ.get('SEND_EMAIL_ADDRESS'),
    to_emails=to_address,
    subject='Monthly {} Sales Report'.format(dept),
    html_content='<strong>Here is the report for the {} department</strong>'.format(dept))
	file_path = str(pathlib.Path().absolute())+'/'+file
	with open(file_path, 'rb') as f:
	    data = f.read()
	    f.close()
	encoded = base64.b64encode(data).decode()
	attachment = Attachment()
	attachment.file_content = FileContent(encoded)
	attachment.file_type = FileType('application/pdf')
	attachment.file_name = FileName(file)
	attachment.disposition = Disposition('attachment')
	attachment.content_id = ContentId('123456')
	message.attachment = attachment
	try:
	    sendgrid_client = SendGridAPIClient(os.environ.get('SENDGRID_API_KEY'))
	    response = sendgrid_client.send(message)
	    print(response.status_code)
	    print(response.body)
	    print(response.headers)
	except Exception as e:
		print(e)

df=pd.read_excel('salesdata.xlsx')
email_dict={
	'Sporting':'sporting@email.com',
	'Toys': 'toys@email.com',
	'Hardware': 'hardware@email.com'
	}

unique_dept=df['Dept'].unique()
for i in unique_dept:
	df_temp=pd.DataFrame(df[df.Dept==i])
	df_temp.to_excel('final_{}.xlsx'.format(i), sheet_name=i)
	send_email(i, email_dict[i], 'final_{}.xlsx'.format(i))


