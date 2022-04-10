# Importing the libraries

import pandas as pd
import win32com.client as client
import streamlit as st
import datetime
import pathlib

#Setting the layout to wide
st.set_page_config(layout='wide')

#Setting the page's title
st.title('Comissions Report')

#------------------ FILE UPLOAD --------------------

#Getting the excel file with the comissions report
comission = st.file_uploader('Upload the comission report', type=['xlsx'])

#After the file is uploaded then:
if comission is not None:
    #Read the file
    comission = pd.read_excel(comission)


#------------------ SIDE BAR --------------------

#Creating the standard date (the 10th of the previous month)
#Getting the current date
today = datetime.date.today()
#Replacing the date with the 1st
first = today.replace(day=1)
#Getting the last month subtracting 1 day from the date
last_month = first - datetime.timedelta(days=1)
#Setting the standard date with the previous month and the 10th
std_date = last_month.replace(day=10)

#When the user uploads a file, then:
if comission is not None:
    #Setting the side bar title
    st.sidebar.title('Pick the dates')
    #Setting the date picker with the standard date as default
    #Credit date
    credit = st.sidebar.date_input('Select the credit date', value=std_date)
    #Deadline date
    deadline = st.sidebar.date_input('Select the deadline date', value=std_date)
    #Setting the side bar title
    st.sidebar.title('Pick the columns')
    #Set the selectbox with the columns to pick the column with the names
    name_column = st.sidebar.selectbox('Select the column with the saler\'s name', comission.columns, index=0)
    #Set the selectbox with the columns to pick the column with the comissions
    comission_column = st.sidebar.selectbox('Select the column with the comissions', comission.columns, index=8)
    #Set the selectbox with the columns to pick the column with the emails
    email_column = st.sidebar.selectbox('Select the column with the emails', comission.columns, index=1)
    #Set a checkbox to check if the user wants a copy to be sent to the manager
    cc = st.sidebar.checkbox('CC manager?', value=True)
    #If is true, then:
    if cc:
        #Set the selectbox with the columns to pick the column with the manager
        manager_column = st.sidebar.selectbox('Select the column with the managers', comission.columns, index=2)
else:
    st.sidebar.title('Waiting file...')


#------------------ SETTING UP THE EMAIL --------------------

#Library function to open the Outlook
otk = client.Dispatch('Outlook.Application')

#If the user uploads a file, then:
if comission is not None:
    #Setting the standard email's subject
    std_subject = 'Invoice details - '+str(std_date.strftime('%B'))
    #Setting in the text input for custom subjects
    subject = st.text_input('Email\'s subject', value=std_subject)

    #Setting the standard email's body
    std_body = 'Hello Mr(s) SELLER_X, \nPlease check the invoice details attached. \n\nReference month: '+str(std_date.strftime('%B'))+'\nComission amount: R$ COMISSION_AMOUNT \nCredit date: '+str(credit)+'\n\nPlease forward all invoice to the Tax Department until 12h of '+str(deadline)+' in order to guarante payment on the right date.\nTax Department: example@email.com'
    #Setting in the text area for custom bodies
    body = st.text_area('Text body to be sent on the email', height=250, value=std_body)


#------------------ SETTING UP FUNCTION --------------------

#Defining the email shooter function passing as parameter the body
def send_emails(body):

#For every unique email in the file, run the loop
    for emails in comission[email_column].unique():
        #Filter the file into a table with the invoices of the current seller
        table = comission[comission[email_column]==emails]
        #Get the name of the seller
        name = str(table[table[email_column] == emails][name_column].unique()[0])
        #Get the name of the manager
        manager = str(table[table[email_column] == emails][manager_column].unique()[0])
        #Setting a file name to save the filtered file
        file_name = name+'_'+str(std_date.strftime('%B'))
        #Saving the filtered file
        table.to_excel(file_name+'.xlsx', index=False)
        #Using the library to fetch the absolute path to the saved file
        path = pathlib.Path(file_name+'.xlsx')
        path = str(path.absolute())

        #Calculating the total comission for the seller and formating it
        comission_amout = str(round(table[comission_column].sum(),2)).replace('.',',')

        #Customizing the body with the seller's name
        body2 = body.replace('SELLER_X', name)
        #Customizing the body with the seller's comission
        body2 = body2.replace('COMISSION_AMOUNT', comission_amout)


        #Creating a new email
        msg = otk.CreateItem(0)
        #Filling the recipient
        msg.To = emails
        #Filling the cc field
        msg.CC = manager
        #Filling the subject field
        msg.Subject = subject
        #Filling the body field
        msg.Body = body2
        #Attaching the filtered file with the absolute path
        msg.Attachments.Add(path)
        #Sending the email
        msg.Send()

    return st.success('Emails sent!')

#Setting the button - if button is pressed it runs the function
if comission is not None:
    if st.button('Send emails'):
        send_emails(body)
        