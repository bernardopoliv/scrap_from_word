import os

import docx
import openpyxl

os.chdir('C:\\Users\BERNARDO\Anaconda3\envs\SCRAP_FROM_WORD')
doc = docx.Document('Residential_Lease_Template_FL22.docx')

fields_and_values = {}

def lookup(paragraph_index: int, runs_index: int, variable_name: str):
    value = doc.paragraphs[paragraph_index].runs[runs_index].text
    fields_and_values.update(
        {
            f'{variable_name}': value
        }
    )
    return value

def get_allparagraph(paragraph_index: int, variable_name: str):
    value = doc.paragraphs[paragraph_index].text
    fields_and_values.update(
        {
            f'{variable_name}': value
        }
    )
    return value

def look_value(paragraph_index: int, runs_index: int):
    value = doc.paragraphs[paragraph_index].runs[runs_index].text

    return value

template_wb = openpyxl.load_workbook('template.xlsx')
template_sheet = template_wb['Sheet1']

def export():
    first_value_row = 2
    column = 1
    for key in fields_and_values.keys():
        template_sheet.cell(row=1, column=column).value = key
        column +=1
        # for value in fields_and_values.values():
        #     try:
        #         for new_value in fields_and_values['f{key}']:
        #             if value == new_value:
        #                 template_sheet.cell(row=first_value_row, column=column).value = value
        #                 first_value_row += 1
        #     except:
        #         pass
        column += 1
        first_value_row = 2
    template_wb.save('C:\\Users\BERNARDO\Anaconda3\envs\SCRAP_FROM_WORD\output.xlsx')



def scan():
    i=0
    for p in doc.paragraphs:
         print(i, p.text)
         i+=1

def find_value(paragraph_index: int):
    for p in range (0,100):
        try:
            print(str(p) + doc.paragraphs[paragraph_index].runs[p].text)
        except:
            pass


def AG_COMMERCIAL():
    fields_and_values = {}
    file = 'Commercial_Lease_Template_FL22.docx'
    doc = docx.Document('Agricultural Crop Share Permit Template FL22.docx')
    template_wb = openpyxl.load_workbook('template.xlsx')
    template_sheet = template_wb['Sheet1']

    #PAGE 1
    lookup(0,0,'PERMIT_TYPE')
    get_allparagraph(2,'AGREEMENT_DATE')
    lookup(6, 3, 'LANDLORD')
    lookup(16, 3, 'TENANT_NAME')
    lookup(28, 3, 'LOT_NUMBER')
    lookup(28, 8, 'PLANNO')

    #PAGE 3
    lookup(57, 1, 'TENANT_LEASE_TERM')
    lookup(57, 4, 'INITIAL_DATE') #TODO: check if it is missing in original variables
    lookup(57, 7, 'EXPIRY_DATE')
    lookup(59, 1, 'TENANT_PREMISE_PURPOSE')

    #PAGE 4
    lookup(65,2,'RENT_EXPIRE_DATE')
    lookup(65, 6, 'YEARLY_RENT_AMOUNT')

    # if 'year' in doc.paragraphs[65].runs[9].text:
    #     fields_and_values['YEARLY_RENT_PAY'] = 'YES'
    # else:
    #     fields_and_values['YEARLY_RENT_PAY'] = 'NO'

    lookup(65, 12, 'MONTH_4_PAYMENT_AGREEMENT')
    lookup(65, 15, 'TENANT_FISCAL_START_DATE')
    lookup(65, 18, 'TENANT_FISCAL_END_DATE')
    lookup(65, 21, 'TENANT_CONFIRMED_PAYMENT_DATE')

    #PAGE 7
    fields_and_values['TENANT_CONIFIRM_PAY_WATER'] = 'YES'
    fields_and_values['TENANT_CONIFIRM_PAY_GAS'] = 'YES'
    fields_and_values['TENANT_CONIFIRM_PAY_TELEPHONE'] = 'YES'
    fields_and_values['TENANT_CONIFIRM_PAY_LIGHT'] = 'YES'
    fields_and_values['TENANT_CONIFIRM_PAY_POWER'] = 'YES'
    fields_and_values['TENANT_CONIFIRM_PAY_HEAT'] = 'YES'
    fields_and_values['TENANT_CONIFIRM_PAY_AIRCONDITIONING'] = 'YES'
    fields_and_values['TENANT_CONIFIRM_PAY_SEWER'] = 'YES'
    fields_and_values['TENANT_CONIFIRM_PAY_GARBAGE'] = 'YES'

    #PAGE 10
    lookup(134,1,'TENANT_LIABILITY_INSURANCE_AMOUNT')

    #PAGE 19
    lookup(233,2,'TENANT_NAME_NOTICE')

    #PAGE 20
    lookup(253, 5,'CHIEF_LAND_COMM_LEASE')
    lookup(255, 5, 'COUNCILLOR1_LAND_COMM_LEASE')
    lookup(257, 5, 'COUNCILLOR2_LAND_COMM_LEASE')
    lookup(260, 6, 'WITNESS1_LAND_COMM_LEASE')

def AG_RESIDENTIAL():
    fields_and_values = {}
    file = 'Residential_Lease_Template_FL22.docx'
    doc = docx.Document('Agricultural Residential_Lease_Template_FL22 Share Permit Template FL22.docx')
    template_wb = openpyxl.load_workbook('template.xlsx')
    template_sheet = template_wb['Sheet1']

    # PAGE 1
    lookup(0, 0, 'PERMIT_TYPE')
    lookup(1, 1, 'RESIDENTIAL_AGREEMENT_DATE')
    lookup(15, 3, 'RESIDENTIAL_TENANT_NAME')

    # PAGE 4
    lookup(62, 4, 'RESIDENTIAL_LEASE_INITIAL_DATE') #TODO: put year together (runs[5)
    lookup(62, 7, 'RESIDENTIAL_LEASE_EXPIRY_DATE') #TODO: put year together (runs[8)

    # PAGE 7
    fields_and_values['RESIDENTIAL_LEASE_TENANT_CONFIRM_PAY_TAXES'] = 'YES'
    fields_and_values['RESIDENTIAL_LEASE_TENANT_CONFIRM_PAY_WATER'] = 'YES'
    fields_and_values['RESIDENTIAL_LEASE_TENANT_CONFIRM_PAY_GAS'] = 'YES'
    fields_and_values['RESIDENTIAL_LEASE_TENANT_CONFIRM_PAY_TELEPHONE'] = 'YES'
    fields_and_values['RESIDENTIAL_LEASE_TENANT_CONFIRM_PAY_LIGHT'] = 'YES'
    fields_and_values['RESIDENTIAL_LEASE_TENANT_CONFIRM_PAY_POWER'] = 'YES'
    fields_and_values['RESIDENTIAL_LEASE_TENANT_CONFIRM_PAY_HEAT'] = 'YES'
    fields_and_values['RESIDENTIAL_LEASE_TENANT_CONFIRM_PAY_AIRCONDITIONING'] = 'YES'
    fields_and_values['RESIDENTIAL_LEASE_TENANT_CONFIRM_PAY_SEWER'] = 'YES'
    fields_and_values['RESIDENTIAL_LEASE_TENANT_CONIFIRM_PAY_GARBAGE'] = 'YES'

    # PAGE 18 TODO: check better this part with client
    lookup(219,2, 'RESIDENTIAL_LEASE_TENANT_NAME_NOTICE')
    lookup(1, 1, 'RESIDENTIAL_LEASE_TENANT_NAME-NOTICE_DELIVERY: MAIL / EMAIL ')
    lookup(15, 3, 'RESIDENTIAL_LEASE_TENANT_NAME_MAILING_ADDRESS')
    lookup(15, 3, 'RESIDENTIAL_LEASE_TENANT_NAME_EMAIL')

    # PAGE 20 TODO: check better this part with client
    lookup(216,2, 'CHIEF_RESIDENTIAL_LEASE')
    lookup(1, 1, 'COUNCILLOR1_RESIDENTIAL_LEASE')
    lookup(15, 3, 'COUNCILLOR2_RESIDENTIAL_LEASE')
    lookup(15, 3, 'WITNESS1_RESIDENTIAL_LEASE')

    # PAGE 20 TODO: check better this part with client
    lookup(216,2, 'TENANT_NAME')
    lookup(1, 1, 'AFFIDAVIT_WITNESS_RESIDENTIAL_LEASE_DATE')
    lookup(15, 3, 'AFFADAVIT_WITNESS_RESIDENTIAL_LEASE_SIGNED: YES/NO')
    lookup(15, 3, 'AFFADAVIT_WITNESS_RESIDENTIAL_LEASE_NOTARY_SIGNED: YES/NO')

def AG_PERMIT():
    fields_and_values = {}

    AG_PERMIT_TYPE = doc.paragraphs[1].text
    AG_PERMIT_TYPE = AG_PERMIT_TYPE[1:]
    fields_and_values.update({'AG_PERMIT_TYPE': AG_PERMIT_TYPE})

    AG_PERM_PRELIM_START_DATE = doc.paragraphs[4].text
    AG_PERM_PRELIM_START_DATE = AG_PERM_PRELIM_START_DATE[35:]
    AG_PERM_PRELIM_START_DATE = AG_PERM_PRELIM_START_DATE.split()

    #
    day = AG_PERM_PRELIM_START_DATE[0][:2]
    month_len = len(AG_PERM_PRELIM_START_DATE[3])
    month = AG_PERM_PRELIM_START_DATE[3][:month_len - 1]
    year = AG_PERM_PRELIM_START_DATE[4][:4]

    AG_PERM_PRELIM_START_DATE = day + ', ' + month + ', ' + year
    fields_and_values.update(
        {
            'AG_PERM_PRELIM_START_DATE': AG_PERM_PRELIM_START_DATE
        }
    )

    lookup(7, 2, "AG_PERM_GRANTOR")
    lookup(9, 3, "AG_PERMIT_GRANTEE")
    lookup(20, 1, "AG_LOCATION")
    lookup(20, 5, "AG_AREA_HECATARES")
    lookup(20, 8, "AG_AREA_ACRES")
    lookup(20, 11, "AG_LAND_USE")

    # PAGE 2
    AG_PERMIT_START_DATE = doc.paragraphs[23].runs[8].text + doc.paragraphs[23].runs[10].text
    fields_and_values.update(
        {
            'AG_PERMIT_START_DATE': AG_PERMIT_START_DATE
        }
    )
    AG_PERMIT_END_DATE = doc.paragraphs[23].runs[15].text + doc.paragraphs[23].runs[17].text
    fields_and_values.update(
        {
            'AG_PERMIT_END_DATE': AG_PERMIT_END_DATE
        }
    )
    AG_PERMIT_COMMENT = doc.paragraphs[31].runs[4].text + doc.paragraphs[31].runs[5].text + \
                        doc.paragraphs[31].runs[6].text
    fields_and_values.update(
        {
            'AG_PERMIT_COMMENT': AG_PERMIT_COMMENT
        }
    )

    # PAGE 3

    lookup(0, 0, 'AG_LAND_CHEMICAL_PRESENT')
    lookup(0, 0, 'AG_LAND_CHEMICAL_TYPE')

    AG_ENVIRONMENTAL_ISSUE = doc.paragraphs[64].text
    AG_ENVIRONMENTAL_ISSUE = AG_ENVIRONMENTAL_ISSUE[4:]  # DEFAULT TO NO
    fields_and_values.update(
        {
            'AG_ENVIRONMENTAL_ISSUE': AG_ENVIRONMENTAL_ISSUE
        }
    )

    lookup(0, 0, 'AG_ENVIRONMENTAL_ISSUE_TYPE')  # DEFAULT TO N/A
    lookup(0, 0, 'AG_LAND_CHEMICAL_PRESENT')

    # PAGE 4
    AG_ADDRESS_NAME = look_value(86, 6)[:-1]
    AG_ADDRESS_NAME = AG_ADDRESS_NAME[:-1]
    fields_and_values.update(
        {
            'AG_ADDRESS_NAME': AG_ADDRESS_NAME
        }
    )

    lookup(87, 7, 'AG_ADDRESS_STREET')
    lookup(88, 7, 'AG_ADDRESS_CITY')
    lookup(88, 10, 'AG_ADDRESS_PROVINCE')
    lookup(89, 7, 'AG_ADDRESS_POSTAL_CODE')
    # PAGE 5 (NO DATA)
    # PAGE 6
    lookup(170, 0, 'AG_SIGNED_CHIEF')
    lookup(171, 0, 'AG_SIGNED_CHIEF_NAME')
    lookup(175, 0, 'AG_SIGNED_COUNCILLOR1')
    lookup(176, 0, 'AG_SIGNED_COUNCILLOR1_NAME')
    lookup(180, 0, 'AG_SIGNED_COUNCILLOR2')
    lookup(181, 0, 'AG_SIGNED_COUNCILLOR2_NAME')
    lookup(141, 0, 'AG_SIGNED_WITNESS')
    lookup(141, 4, 'AG_SIGNED_PERMITTEE_NAME')
    lookup(181, 0, 'AG_SIGNED_COUNCILLOR2_NAME')
    lookup(181, 0, 'AG_SIGNED_COUNCILLOR2_NAME')
