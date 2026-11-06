import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 1 HERE
    
while True:
    dir_string = single_input("Please input the desired directory file path, no quotations")
    directory = dir_string + '\\'
    
    try:
        dir_files = os.listdir(directory)
        if 'ProductData.xlsx' in dir_files and 'MasterInvoice.pdf' in dir_files and 'Order1.xlsx' in dir_files and 'Order2.xlsx' in dir_files and 'Order3.xlsx' in dir_files and 'Order4.xlsx' in dir_files:
            show_message('valid',f"Valid directory. Your directory is {directory}")
            break
        else:
            show_message('error','An error occurred. Please re-enter your directory path.')
    except:
        show_message('error_1','Not a valid directory. Please try again.')

print(directory)

# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}

# product descriptions
product_data_str = 'ProductData.xlsx'
Pd_desc = pd.read_excel(directory+product_data_str,sheet_name='Descriptions')

for index,row in Pd_desc.iterrows():
    description = row['Description']
    unit_number = row['Unit number']
    product_descriptions[unit_number] = description

# product prices
product_data_str = 'ProductData.xlsx'
Pd_price = pd.read_excel(directory+product_data_str,sheet_name='Prices')

for index,row in Pd_price.iterrows():
    price_per_unit = str(row['Price per unit'])
    unit_number = row['Unit number']
    product_prices[unit_number] = price_per_unit

# results
product_descriptions
product_prices

# INSERT ANSWER FOR QUESTION 3 HERE

order_data = {}

def get_order_data(invoice_num,file,directory):
    
    template = openpyxl.load_workbook(directory+file)
    selected_sheet = template['OrderForm']

    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    month = selected_sheet['C9'].value.month
    day = selected_sheet['C9'].value.day
    year = selected_sheet['C9'].value.year
    date = (f'{month}/{day}/{year}')

    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    address = selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zip = selected_sheet['C19'].value
    order_data['CustomerAddress1'] = (f"{address}")
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = (f"{city}, {state} {zip}")

    i = 0
    total = 0
    while 0 <= i <= 5:
        n = 24+i
        if str(product_prices.get(selected_sheet['B'+str(n)].value)) == 'None':
            order_data[('Price'+str(i))] = ''
            order_data[('Quantity'+str(i))] = ''
            order_data[('Amount'+str(i))] = ''
            order_data[('Description'+str(i))] = ''
        else:
            order_data[('Price'+str(i))] = f'${float(product_prices.get(selected_sheet['B'+str(n)].value)):,.2f}'
            order_data[('Quantity'+str(i))] = str(selected_sheet['C'+str(n)].value)
            price = float(str(product_prices.get(selected_sheet['B'+str(n)].value)))
            quantity = int(selected_sheet['C'+str(n)].value)
            amount = price*quantity
            total = total + amount
            order_data[('Amount'+str(i))] = f'${price*quantity:,.2f}'
            order_data[('Description'+str(i))] = str(product_descriptions.get(selected_sheet['B'+str(n)].value))
        i = i + 1

    totalstr = f'${total:,.2f}'
    order_data['Total'] = totalstr
    
    
    return order_data, name, email, date

order_data, name, email, date = get_order_data(5818, 'Order1.xlsx', directory)
print(order_data)
print(name)
print(email)
print(date)


# INSERT ANSWER FOR QUESTION 4 HERE
pdf_name = 'MasterInvoice.pdf'

invoice_num = 8515
def create_invoice(invoice_num,order_data,directory):
    invoice_num = int(invoice_num)
    create_pdf(directory+pdf_name, f'Invoice {invoice_num}.pdf', order_data)


create_invoice(invoice_num,order_data,directory)

# INSERT ANSWER FOR QUESTION 5 HERE

gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):
    gmail_user = email
    
    
    # List of recipients

    recipients = gmail_user

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    E-mail: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file1 = directory+f'Invoice {invoice_num}.pdf'

    messsage = add_attachment(path_to_file1, message)

            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(invoice_num,name,email,date,directory)
    
# INSERT ANSWER FOR QUESTION 6 HERE
filenames1 = os.listdir()
filenames = sorted(filenames1)

invoice_num_start = 8515
loop = 0

for filename in filenames:
    if 'Order' in filename and '.xlsx' in filename:
        invoice_num = invoice_num_start + loop
        order_data,name,email,date = get_order_data(invoice_num,filename,directory)
        create_invoice(invoice_num,order_data,directory)
        send_email(invoice_num,name,email,date,directory)
        loop = loop + 1


