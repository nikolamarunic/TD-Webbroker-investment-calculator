#Constants for sheet locations.
#Columns
CDN_BOND = 'B'
CDN_INDX = 'C'
USA_INDX = 'D'
INTL_INDX = 'E'
holding_locations = [CDN_BOND, CDN_INDX, USA_INDX, INTL_INDX]

#Rows
CAD_CASH = '3'
CAD_TFSA = '4'

#function for format the new excel sheet
def format_sheet(sheet, name):
    sheet['A2'] = 'Name'
    sheet['A3'] = 'CAD Cash'
    sheet['A4'] = 'CAD TFSA'
    sheet['A6'] = 'Cash Added'
    sheet['A7'] = 'TFSA Added'
    
    sheet['B2'] = 'Canadian Index'
    sheet['D2'] = 'Canadian Bond Index'
    sheet['C2'] = 'American Index'
    sheet['E2'] = 'Int\'l Index'
    sheet['A1'] = name
    sheet['F1'] = 'TOTAL'
