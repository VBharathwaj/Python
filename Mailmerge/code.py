from mailmerge import MailMerge
from lxml import etree

from datetime import date

template = "template.docx"
document = MailMerge(template)
print(document.get_merge_fields())

document.merge(
	invoice_number='1',
	invoice_date='{:%d-%b-%Y}'.format(date.today()),
    customer_name='Wolverine',
    customer_address_line_1='1, Downing Street',
    customer_address_line_2='San Francisco',
    customer_state_and_zip='California 100 001',
    customer_phone_number='9876546789',
    company_name='Stark Industries',
    company_address_line1='1, Wall Street',
    company_address_line_2='Houston',
    company_state_and_zip='Texas 100 002',
	company_phone_number='8756785989',
	delivery_charge='$150',
	total='$9150',
	benificiary_name='Tony Stark',
	benificiary_account_number='1234 5678 9101 1121',
	bank_name_and_address='Soft Bank, New York',
	ifsc_code='SB000921',
	upi_handle='@stark_sb',
	contact_name_1='Tony',
	contact_number_1='9876546789',
	contact_name_2='Stark',
	contact_number_2='8756789879')
	
purchase = [{
	's_no':'1',
	'product_name':'Ark reactor',
	'quantity':'3',
	'rate':'$2000',
	'amount':'$6000'
	},{
	's_no':'2',
	'product_name':'Mini Ark reactor',
	'quantity':'2',
	'rate':'$1500',
	'amount':'$3000'
	}];

document.merge_rows('s_no',purchase)
	
document.write('result.docx')