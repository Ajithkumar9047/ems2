import pyodbc
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
import tempfile
import logging
current_date = datetime.now().replace(hour=18, minute=0, second=0, microsecond=0).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
yesterday_date = (datetime.now() - timedelta(days=1)).replace(hour=12, minute=0, second=0, microsecond=0).strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
#current_date ='2024-01-18 00:00:00.000'
#yesterday_date ='2024-10-22 00:00:00.000'
ccMail=['mohit.lalwani@unicorntechmedia.com', 'kirti.jindal@unicorntechmedia.com', 'nisha.singh@unicorntechmedia.com', 'ankur.kataria@unicorntechmedia.com', 'arun.reddy@newgendigital.com','Neha.Parasher@tvsmotor.com', 
'prasanth@newgendigital.com', 'priyadharshini@newgendigital.com', 'punitha@newgendigital.com', 'kartik.shah@unicorntechmedia.com', 'somesh@unicorntechmedia.com']
logging.basicConfig(
    filename='lmsapp.log',
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s]: %(message)s',
)

query1 = f"""SELECT *
FROM tvs_91Wheels_leads
WHERE 
    create_date BETWEEN '{yesterday_date}' AND '{current_date}'
    AND ISNULL(ems_server_response, '') = ''
    AND ISNULL(dms_response, '') = ''
    AND ISNULL(repush_ems, 0) = 0
    AND id NOT IN (
        SELECT crm.CRM_INTERNET_ENQUIRY_ID
        FROM tb_tvs_common_crm_api_leads crm
        JOIN tvs_91Wheels_leads web ON crm.CRM_INTERNET_ENQUIRY_ID = web.id
        WHERE 
            web.create_date BETWEEN '{yesterday_date}' AND '{current_date}'
            AND ISNULL(web.ems_server_response, '') = ''
            AND ISNULL(web.dms_response, '') = ''
    );"""
query2 = f""" select id as LeadId,INTERNET_ENQUIRY_ID,ems_server_response,name,create_date_get
            from tvs_91Wheels_leads  where create_date BETWEEN '{yesterday_date}' AND '{current_date}'  
            and isnull(repush_ems,0)=1 ORDER BY 
            create_date_get ASC"""
query3 = f"""  select web.id as LeadId,web.INTERNET_ENQUIRY_ID,crm.ems_server_response,web.name,crm.dms_response,web.create_date
            from tb_tvs_common_crm_api_leads crm join tvs_91Wheels_leads web
            on crm.CRM_INTERNET_ENQUIRY_ID=web.id
            where web.create_date BETWEEN '{yesterday_date}' AND '{current_date}' 
            and ISNULL(web.ems_server_response,'')='' 
            and  ISNULL(web.dms_response,'')='' ORDER BY 
            web.create_date ASC;"""
