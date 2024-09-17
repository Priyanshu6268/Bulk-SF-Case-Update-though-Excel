import pandas as pd
import win32com.client
import streamlit as st

# Function to send email
def send_email(subject, body, cc_address):
    ol = win32com.client.Dispatch('Outlook.Application')
    to_address=f'support.services@nokia.com'
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = subject
    newmail.To = to_address
    newmail.CC = cc_address  # Set the Owner mail in CC
    newmail.Body = body
    newmail.Send()
    return f"Email sent to {to_address} with CC to {cc_address} and subject '{subject}'"

# Streamlit interface
st.title('Bulk SF-Case Update though Excel')

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    # Load the Excel file into a DataFrame
    df = pd.read_excel(uploaded_file)

    # Display the content of the DataFrame
    st.write("Data from Excel file:")
    st.write(df)

    if st.button('Send Emails'):
        results = []
        for _, row in df.iterrows():
            # case_no = row['Case no']
            # tid = row['Tid']
            owner_mail = row['OwnerEmail'].strip()  # Remove any extra spaces
            update = row['Update']
            
            subject = row['Subject']
            body = update

            # Send email from 'priyanshu.kumar_saw@nokia.com' and CC the owner mail
            # from_address = 'priyanshu.kumar_saw@nokia.com'
            cc_address = owner_mail

            result = send_email(subject, body, cc_address)
            results.append(result)
        
        # Display the results
        st.write("Email sending results:")
        for result in results:
            st.write(result)
