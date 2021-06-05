'''
Agentbox is Australia's #1 Real Estate CRM, it offers
agencies a complete cloud-based software solution for
Sales and Property Management teams. This project aims
to read all customer information through Agentbox's API
and store it in an Excel table. The information is stored
in the form of Json.
'''
import requests
import json
headers = {
    'Accept': 'application/json',
    'X-Client-ID': 'client_ID',
    'X-API-Key': 'API_key',
}

params = (
    ('include', 'searchRequirements'),
    ('version', '2'),
)
final = []
for i in range(200):
    l=[]
    a = str(i)
    url = 'https://api.agentboxcrm.com.au/contacts/'+ a
    response = requests.get(url, headers=headers, params=params)
    jsondata = ''
    for a in response:
        str_data = a.decode()
        jsondata+=str_data
    text = json.loads(jsondata)
    try:
        l.append(text['response']['contact']['id'])
        l.append(text['response']['contact']['clientRef'])
        l.append(text['response']['contact']['type'])
        l.append(text['response']['contact']['status'])
        l.append(text['response']['contact']['title'])
        l.append(text['response']['contact']['firstName'])
        l.append(text['response']['contact']['lastName'])
        l.append(text['response']['contact']['website'])
        l.append(text['response']['contact']['salutation'])
        l.append(text['response']['contact']['customSalutation'])
        l.append(text['response']['contact']['addressTo'])
        l.append(text['response']['contact']['legalName'])
        l.append(text['response']['contact']['companyName'])
        l.append(text['response']['contact']['jobTitle'])
        l.append(text['response']['contact']['email'])
        l.append(text['response']['contact']['mobile'])
        l.append(text['response']['contact']['homePhone'])
        l.append(text['response']['contact']['workPhone'])
        l.append(text['response']['contact']['fax'])
        l.append(text['response']['contact']['streetAddress']['address'])
        l.append(text['response']['contact']['streetAddress']['suburb'])
        l.append(text['response']['contact']['streetAddress']['state'])
        l.append(text['response']['contact']['streetAddress']['country'])
        l.append(text['response']['contact']['streetAddress']['postcode'])
        l.append(text['response']['contact']['postalAddress']['address'])
        l.append(text['response']['contact']['postalAddress']['suburb'])
        l.append(text['response']['contact']['postalAddress']['state'])
        l.append(text['response']['contact']['postalAddress']['country'])
        l.append(text['response']['contact']['postalAddress']['postcode'])
        l.append(text['response']['contact']['letterAddressBlock'])
        l.append(text['response']['contact']['prefContactMethod'])
        l.append(text['response']['contact']['source'])
        l.append(text['response']['contact']['comments'])
        l.append(text['response']['contact']['firstCreated'])
        l.append(text['response']['contact']['lastModified'])
        l.append(text['response']['contact']['lastContacted'])
        l.append(text['response']['contact']['attachedRelatedStaffMembers'][0]['role'])
        l.append(text['response']['contact']['attachedRelatedStaffMembers'][0]['id'])
        l.append(text['response']['contact']['contactClasses'][0]['id'])
        l.append(text['response']['contact']['contactClasses'][0]['name'])
        l.append(text['response']['contact']['contactClasses'][0]['type'])
        l.append(text['response']['contact']['keyDates'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['communicationRestrictions']['doNotCall'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['communicationRestrictions']['doNotSMS'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['communicationRestrictions']['doNotEmail'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['communicationRestrictions']['doNotMail'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['id'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['name'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['contactId'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['listingType'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['propertyType'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['propertyCategories'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['price'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['bedrooms'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['bathrooms'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['parking'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['landArea'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['buildingArea'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['suburbs'][0]['name'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['suburbs'][0]['postcode'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['suburbs'][0]['state'])
    except:
        l.append('not given')
    try:
        l.append(text['response']['contact']['searchRequirements'][0]['suburbs'][0]['country'])
    except:
        l.append('not given')
    final.append(l)
print(final[-1])

import xlsxwriter
import pdfplumber
import os

# Create Excel file
xl = xlsxwriter.Workbook(r'C:\Users\Eleni\agentbox6.xlsx')
# add sheet
sheet = xl.add_worksheet('sheet1')
sheet.write_string("A1", "contactID")
sheet.write_string("B1", "clientRef")
sheet.write_string("C1", "type")
sheet.write_string("D1", "status")
sheet.write_string("E1", "title")
sheet.write_string("F1", "first name")
sheet.write_string("G1", "last name")
sheet.write_string("H1", "website")
sheet.write_string("I1", "salutation")
sheet.write_string("J1", "custom salutation")
sheet.write_string("K1", "address to")
sheet.write_string("L1", "legal name")
sheet.write_string("M1", "company name")
sheet.write_string("N1", "job title")
sheet.write_string("O1", "email")
sheet.write_string("P1", "mobile")
sheet.write_string("Q1", "home phone")
sheet.write_string("R1", "work phone")
sheet.write_string("S1", "fax")
sheet.write_string("T1", "street address")
sheet.write_string("U1", "suburb")
sheet.write_string("V1", "state")
sheet.write_string("W1", "country")
sheet.write_string("X1", "postcode")
sheet.write_string("Y1", "postal address")
sheet.write_string("Z1", "suburb")
sheet.write_string("AA1", "state")
sheet.write_string("AB1", "country")
sheet.write_string("AC1", "postcode")
sheet.write_string("AD1", "letteraddressblock")
sheet.write_string("AE1", "prefer contact method")
sheet.write_string("AF1", "source")
sheet.write_string("AG1", "comments")
sheet.write_string("AH1", "firstcreated")
sheet.write_string("AI1", "lastmodified")
sheet.write_string("AJ1", "lastcontacted")
sheet.write_string("AK1", "related staff role")
sheet.write_string("AL1", "related staff id")
sheet.write_string("AM1", "contact class id")
sheet.write_string("AN1", "contact class name")
sheet.write_string("AO1", "contact class type")
sheet.write_string("AP1", "key dates")
sheet.write_string("AQ1", "do not call")
sheet.write_string("AR1", "do not sms")
sheet.write_string("AS1", "do not email")
sheet.write_string("AT1", "do not mail")
sheet.write_string("AU1", "search id")
sheet.write_string("AV1", "search name")
sheet.write_string("AW1", "search contact id")
sheet.write_string("AX1", "listing type")
sheet.write_string("AY1", "property type")
sheet.write_string("AZ1", "property categories")
sheet.write_string("BA1", "price")
sheet.write_string("BB1", "bedrooms")
sheet.write_string("BC1", "bathrooms")
sheet.write_string("BD1", "parking")
sheet.write_string("BE1", "landarea")
sheet.write_string("BF1", "buildingarea")
sheet.write_string("BG1", "suburbs name")
sheet.write_string("BH1", "suburbs postcode")
sheet.write_string("BI1", "suburbs state")
sheet.write_string("BJ1", "suburbs country")

x = 0
for i in range(len(final)):
    x += 1
    y = x + 1
    sheet.write_string("A%d" % y, "%s" % final[i][0])
    sheet.write_string("B%d" % y, "%s" % final[i][1])
    sheet.write_string("C%d" % y, "%s" % final[i][2])
    sheet.write_string("D%d" % y, "%s" % final[i][3])
    sheet.write_string("E%d" % y, "%s" % final[i][4])
    sheet.write_string("F%d" % y, "%s" % final[i][5])
    sheet.write_string("G%d" % y, "%s" % final[i][6])
    sheet.write_string("H%d" % y, "%s" % final[i][7])
    sheet.write_string("I%d" % y, "%s" % final[i][8])
    sheet.write_string("J%d" % y, "%s" % final[i][9])
    sheet.write_string("K%d" % y, "%s" % final[i][10])
    sheet.write_string("L%d" % y, "%s" % final[i][11])
    sheet.write_string("M%d" % y, "%s" % final[i][12])
    sheet.write_string("N%d" % y, "%s" % final[i][13])
    sheet.write_string("O%d" % y, "%s" % final[i][14])
    sheet.write_string("P%d" % y, "%s" % final[i][15])
    sheet.write_string("Q%d" % y, "%s" % final[i][16])
    sheet.write_string("R%d" % y, "%s" % final[i][17])
    sheet.write_string("S%d" % y, "%s" % final[i][18])
    sheet.write_string("T%d" % y, "%s" % final[i][19])
    sheet.write_string("U%d" % y, "%s" % final[i][20])
    sheet.write_string("V%d" % y, "%s" % final[i][21])
    sheet.write_string("W%d" % y, "%s" % final[i][22])
    sheet.write_string("X%d" % y, "%s" % final[i][23])
    sheet.write_string("Y%d" % y, "%s" % final[i][24])
    sheet.write_string("Z%d" % y, "%s" % final[i][25])
    sheet.write_string("AA%d" % y, "%s" % final[i][26])
    sheet.write_string("AB%d" % y, "%s" % final[i][27])
    sheet.write_string("AC%d" % y, "%s" % final[i][28])
    sheet.write_string("AD%d" % y, "%s" % final[i][29])
    sheet.write_string("AE%d" % y, "%s" % final[i][30])
    sheet.write_string("AF%d" % y, "%s" % final[i][31])
    sheet.write_string("AG%d" % y, "%s" % final[i][32])
    sheet.write_string("AH%d" % y, "%s" % final[i][33])
    sheet.write_string("AI%d" % y, "%s" % final[i][34])
    sheet.write_string("AJ%d" % y, "%s" % final[i][35])
    sheet.write_string("AK%d" % y, "%s" % final[i][36])
    sheet.write_string("AL%d" % y, "%s" % final[i][37])
    sheet.write_string("AM%d" % y, "%s" % final[i][38])
    sheet.write_string("AN%d" % y, "%s" % final[i][39])
    sheet.write_string("AO%d" % y, "%s" % final[i][40])
    sheet.write_string("AP%d" % y, "%s" % final[i][41])
    sheet.write_string("AQ%d" % y, "%s" % final[i][42])
    sheet.write_string("AR%d" % y, "%s" % final[i][43])
    sheet.write_string("AS%d" % y, "%s" % final[i][44])
    sheet.write_string("AT%d" % y, "%s" % final[i][45])
    sheet.write_string("AU%d" % y, "%s" % final[i][46])
    sheet.write_string("AV%d" % y, "%s" % final[i][47])
    sheet.write_string("AW%d" % y, "%s" % final[i][48])
    sheet.write_string("AX%d" % y, "%s" % final[i][49])
    sheet.write_string("AY%d" % y, "%s" % final[i][50])
    sheet.write_string("AZ%d" % y, "%s" % final[i][51])
    sheet.write_string("BA%d" % y, "%s" % final[i][52])
    sheet.write_string("BB%d" % y, "%s" % final[i][53])
    sheet.write_string("BC%d" % y, "%s" % final[i][54])
    sheet.write_string("BD%d" % y, "%s" % final[i][55])
    sheet.write_string("BE%d" % y, "%s" % final[i][56])
    sheet.write_string("BF%d" % y, "%s" % final[i][57])
    sheet.write_string("BG%d" % y, "%s" % final[i][58])
    sheet.write_string("BH%d" % y, "%s" % final[i][59])
    sheet.write_string("BI%d" % y, "%s" % final[i][60])
    sheet.write_string("BJ%d" % y, "%s" % final[i][61])

xl.close()
