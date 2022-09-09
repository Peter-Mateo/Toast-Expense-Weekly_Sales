import datetime
import os
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
import win32com.client as client
from openpyxl import Workbook, load_workbook
from openpyxl.styles import (Alignment, Border, Font, PatternFill, Protection,
                             Side)

import models.xls_to_xlsx as xls

# Converts the xls file to xlsx format
xls.converter()

# Creates the new Workbook
main_workbook = load_workbook("C:\\Users\\12392\\Desktop\\Github\\Toast-Expense-Weekly_Sales\\template\\template.xlsx")
main_sheet = main_workbook.active
main_sheet.title = "Date" # Fill in the date with the last date


"""
#Creates the styling of the sheet 
# Column width 
main_sheet.column_dimensions['A'].width = 19.58
main_sheet.column_dimensions['B'].width = 11.15
main_sheet.column_dimensions['C'].width = 11.15
main_sheet.column_dimensions['D'].width = 11.15
main_sheet.column_dimensions['E'].width = 11.15
main_sheet.column_dimensions['F'].width = 11.15
main_sheet.column_dimensions['G'].width = 11.15
main_sheet.column_dimensions['H'].width = 11.15
main_sheet.column_dimensions['J'].width = 11.15
main_sheet.column_dimensions['I'].width = 1.42
main_sheet.column_dimensions['K'].width = 1.42
main_sheet.column_dimensions['M'].width = 11.42
main_sheet.column_dimensions['L'].width = 11.42
main_sheet.column_dimensions['N'].width = 6.85
main_sheet.column_dimensions['O'].width = 6.15
main_sheet.column_dimensions['P'].width = 17.42
# Weekly Sales Acct
main_sheet.merge_cells('A1:B1')
main_sheet['A1'].font = Font(bold=True, size=14)
main_sheet['A1'] = 'WEEKLY SALES ACCT'
# Week Starting
main_sheet.merge_cells('E1:F1')
main_sheet['E1'].font = Font(bold=True, size=12)
main_sheet['E1'] = 'Week Starting'
# Bold and italic styling
font_bold = Font(bold=True, size = 10)
font_italic = Font(italic=True, size = 9)
# Column A Header
main_sheet['A7'] = 'total deposit'
main_sheet['A7'].font = font_italic
main_sheet['A8'] = 'GIFT CARDS (SOLD)'
main_sheet['A8'].font = font_bold
main_sheet['A9'] = 'GIFT CARDS (REDEEMED)'
main_sheet['A9'].font = font_bold
main_sheet['A10'] = 'CC FEES'
main_sheet['A10'].font = font_bold
main_sheet['A11'] = 'CR. CARD DEPOSIT'
main_sheet['A11'].font = font_bold
main_sheet['A12'] = 'CASH SALES'
main_sheet['A12'].font = font_bold
main_sheet['A13'] = 'UBER EATS'
main_sheet['A13'].font = font_bold
main_sheet['A14'] = 'TIPS PAID'
main_sheet['A14'].font = font_bold
main_sheet['A15'] = 'S-FOOD'
main_sheet['A15'].font = font_bold
main_sheet['A16'] = 'S-BAR'
main_sheet['A16'].font = font_bold
main_sheet['A17'] = 'S-MISC'
main_sheet['A17'].font = font_bold
main_sheet['A18'] = 'COMP PROMO'
main_sheet['A18'].font = font_bold
main_sheet['A19'] = 'MGR DISCOUNT'
main_sheet['A19'].font = font_bold
main_sheet['A20'] = 'PROMO'
main_sheet['A20'].font = font_bold
main_sheet['A21'] = 'SERV.DISCOUNT'
main_sheet['A21'].font = font_bold
main_sheet['A22'] = 'QSA'
main_sheet['A22'].font = font_bold
main_sheet['A23'] = 'SALES TAX'
main_sheet['A23'].font = font_bold
main_sheet['A25'] = 'BALANCE'
main_sheet['A25'].font = font_italic
main_sheet['A25'].alignment = Alignment(horizontal = 'right')
main_sheet['A27'] = 'CASH'
main_sheet['A27'].font = font_bold

# MON - Sun
main_sheet['B4'] = 'MON'
main_sheet['C4'] = 'TUE'
main_sheet['D4'] = 'WED'
main_sheet['E4'] = 'THU'
main_sheet['F4'] = 'FRI'
main_sheet['G4'] = 'SAT'
main_sheet['H4'] = 'SUN'

# Entry Box
#Fix Center Alignment 
main_sheet.merge_cells('L5:M5')
main_sheet['L5'] = 'Entry'
main_sheet['L5'].font = Font(bold=True, size= 11)
main_sheet['L5'].alignment = Alignment(horizontal = 'center')
main_sheet['L6'] = 'DR'
main_sheet['L6'].alignment = Alignment(horizontal = 'right')
main_sheet['M6'] = 'CR'
main_sheet['M6'].alignment = Alignment(horizontal = 'right')
main_sheet['N24'] = 'MON'
main_sheet['N24'].alignment = Alignment(horizontal = 'center')
main_sheet['N25'] = 'TUE'
main_sheet['N25'].alignment = Alignment(horizontal = 'center')
main_sheet['N26'] = 'WED'
main_sheet['N26'].alignment = Alignment(horizontal = 'center')
main_sheet['N27'] = 'THU'
main_sheet['N27'].alignment = Alignment(horizontal = 'center')
main_sheet['N28'] = 'FRI'
main_sheet['N28'].alignment = Alignment(horizontal = 'center')
main_sheet['N29'] = 'SAT'
main_sheet['N29'].alignment = Alignment(horizontal = 'center')
main_sheet['N30'] = 'SUN'
main_sheet['N30'].alignment = Alignment(horizontal = 'center')

pcolumn = ['GIFT CARDS','GIFT CARDS','CC FEES','IBERIA','RESTAURANT CASH','UBER EATS','RESTAURANT CASH','SALES - FOOD','SALES -BAR','SALES -MISC','COMP PROMO','MGR DISCOUNT','PROMO','SERV. DISCOUNT','QSA','SALES TAXES PAYABLE','IBERIA','IBERIA','IBERIA','IBERIA','IBERIA','IBERIA','IBERIA']
ocolumn = [21000,21000,40200,10010,10030,10060,10030,40001,40002,40003,40011,40012,40013,40014,40015,28000,10010,10010,10010,10010,10010,10010,10010]

entry_o = 8
for i in range(len(pcolumn)):
    main_sheet.cell(entry_o,column = 16, value=pcolumn[i])
    main_sheet.cell(entry_o,column = 15, value=ocolumn[i])
    entry_o += 1
# Entry Styling 
# Column L
main_sheet['L5'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L6'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L7'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L8'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L9'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L10'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L11'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L12'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L13'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L14'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L15'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L16'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L17'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L18'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L19'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L20'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L21'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L22'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L23'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L24'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L25'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L26'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L27'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L28'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L29'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['L30'].fill = PatternFill('solid', start_color='D9D9D9')
# Column M
main_sheet['M5'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M6'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M7'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M8'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M9'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M10'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M11'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M12'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M13'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M14'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M15'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M16'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M17'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M18'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M19'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M20'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M21'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M22'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M23'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M24'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M25'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M26'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M27'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M28'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M29'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['M30'].fill = PatternFill('solid', start_color='D9D9D9')
# Column N
main_sheet['N5'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N6'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N7'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N8'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N9'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N10'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N11'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N12'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N13'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N14'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N15'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N16'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N17'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N18'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N19'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N20'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N21'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N22'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N23'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N24'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N25'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N26'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N27'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N28'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N29'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['N30'].fill = PatternFill('solid', start_color='D9D9D9')
# Column O
main_sheet['O5'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O6'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O7'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O8'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O9'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O10'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O11'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O12'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O13'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O14'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O15'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O16'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O17'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O18'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O19'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O20'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O21'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O22'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O23'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O24'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O25'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O26'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O27'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O28'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O29'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O30'].fill = PatternFill('solid', start_color='D9D9D9')
# Column P
main_sheet['P5'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P6'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P7'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P8'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P9'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P10'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P11'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P12'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P13'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P14'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P15'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P16'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P17'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P18'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P19'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P20'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P21'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P22'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P23'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P24'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P25'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P26'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P27'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P28'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P29'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P30'].fill = PatternFill('solid', start_color='D9D9D9')
# Font
size = Font(size = 10)
main_sheet['O8'].font = size
main_sheet['O9'].font = size
main_sheet['O10'].font = size
main_sheet['O11'].font = size
main_sheet['O12'].font = size
main_sheet['O13'].font = size
main_sheet['O14'].font = size
main_sheet['O15'].font = size
main_sheet['O16'].font = size
main_sheet['O17'].font = size
main_sheet['O18'].font = size
main_sheet['O19'].font = size
main_sheet['O20'].font = size
main_sheet['O21'].font = size
main_sheet['O22'].font = size
main_sheet['O23'].font = size
main_sheet['O24'].font = size
main_sheet['O25'].font = size
main_sheet['O26'].font = size
main_sheet['O27'].font = size 
main_sheet['O28'].font = size
main_sheet['O29'].font = size
main_sheet['O30'].font = size
# P Font Size
main_sheet['P8'].font = size
main_sheet['P9'].font = size
main_sheet['P10'].font = size
main_sheet['P11'].font = size
main_sheet['P12'].font = size
main_sheet['P13'].font = size
main_sheet['P14'].font = size
main_sheet['P15'].font = size
main_sheet['P16'].font = size
main_sheet['P17'].font = size
main_sheet['P18'].font = size
main_sheet['P19'].font = size
main_sheet['P20'].font = size
main_sheet['P21'].font = size
main_sheet['P22'].font = size
main_sheet['P23'].font = size
main_sheet['P24'].font = size
main_sheet['P25'].font = size
main_sheet['P26'].font = size
main_sheet['P27'].font = size 
main_sheet['P28'].font = size
main_sheet['P29'].font = size
main_sheet['P30'].font = size
# Date JR
main_sheet['O1'] = 'Date:'
main_sheet['O1'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['O1'].font = Font(bold=True, size=14)
main_sheet['O2'] = 'Jr:'
main_sheet['O2'].font = Font(bold=True, size=14)
main_sheet['O2'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P1'].fill = PatternFill('solid', start_color='D9D9D9')
main_sheet['P2'].fill = PatternFill('solid', start_color='D9D9D9')
"""

dir = os.getcwd() + '/new version/'

for files in os.listdir(dir):
    wb = load_workbook(dir + files)
    # Opens the first sheet in the workbook
    ws = wb['Summary']

    """ Start of Everything """
    # Locates Payment Summary 
    payment_summary = xls.search_value_in_column(ws, 'Payment Summary')
    # Locates Total in Payment Summary row
    payment_summary_total = xls.search_value_in_row(ws, 'Total', payment_summary[1])
    # Locates Credit in Payment Summary 
    credit_location = xls.search_value_in_column(ws, 'Credit', 'B')
    # Total deposit
    x = credit_location[1]
    y = payment_summary_total[0]
    z = y + str(x)
    total_deposit = ws[z].value
    # Loops through each of the files in order fill in the data
    if main_sheet['B7'].value == None:
        main_sheet['B7'] = total_deposit
    elif main_sheet['C7'].value == None:
        main_sheet['C7'] = total_deposit
    elif main_sheet['D7'].value == None:
        main_sheet['D7'] = total_deposit
    elif main_sheet['E7'].value == None:
        main_sheet['E7'] = total_deposit
    elif main_sheet['F7'].value == None:
        main_sheet['F7'] = total_deposit
    elif main_sheet['G7'].value == None:
        main_sheet['G7'] = total_deposit
    elif main_sheet['H7'].value == None:
        main_sheet['H7'] = total_deposit

    # Gift Cards SOlD 
    sales_summary = xls.search_value_in_column(ws, 'Sales Summary')
    deferred = xls.search_value_in_row(ws, 'Deferred', sales_summary[1])
    x = int(sales_summary[1]) + 1
    y = deferred[0]
    z = y + str(x)
    gift_sold_value = ws[z].value
    gift_sold = gift_sold_value.strip('$')
    # Loops through each of the files in order fill in the data
    if main_sheet['B8'].value == None:
        main_sheet['B8'] = int(gift_sold)
    elif main_sheet['C8'].value == None:
        main_sheet['C8'] = int(gift_sold)
    elif main_sheet['D8'].value == None:
        main_sheet['D8'] = int(gift_sold)
    elif main_sheet['E8'].value == None:
        main_sheet['E8'] = int(gift_sold)
    elif main_sheet['F8'].value == None:
        main_sheet['F8'] = int(gift_sold)
    elif main_sheet['G8'].value == None:
        main_sheet['G8'] = int(gift_sold)
    elif main_sheet['H8'].value == None:
        main_sheet['H8'] = int(gift_sold)

    # Gift Cards REDEEMED
    gift_card_coords = xls.search_value_in_column(ws, 'Gift Card', 'B')
    x = gift_card_coords[1]
    y = payment_summary_total[0]
    z = y + str(x)
    gift_card_total = ws[z].value
    # Loops through each of the files in order fill in the data
    if main_sheet['B9'].value == None:
        main_sheet['B9'] = gift_card_total
    elif main_sheet['C9'].value == None:
        main_sheet['C9'] = gift_card_total
    elif main_sheet['D9'].value == None:
        main_sheet['D9'] = gift_card_total
    elif main_sheet['E9'].value == None:
        main_sheet['E9'] = gift_card_total
    elif main_sheet['F9'].value == None:
        main_sheet['F9'] = gift_card_total
    elif main_sheet['G9'].value == None:
        main_sheet['G9'] = gift_card_total
    elif main_sheet['H9'].value == None:
        main_sheet['H9'] = gift_card_total

    # CASH SALES
    cash_location = xls.search_value_in_column(ws, 'Cash', 'B')
    x = cash_location[1]
    y = payment_summary_total[0]
    z = y + str(x)
    cash_total = ws[z].value
    # Loops through each of the files in order fill in the data
    if main_sheet['B12'].value == None:
        main_sheet['B12'] = cash_total
    elif main_sheet['C12'].value == None:
        main_sheet['C12'] = cash_total
    elif main_sheet['D12'].value == None:
        main_sheet['D12'] = cash_total
    elif main_sheet['E12'].value == None:
        main_sheet['E12'] = cash_total
    elif main_sheet['F12'].value == None:
        main_sheet['F12'] = cash_total
    elif main_sheet['G12'].value == None:
        main_sheet['G12'] = cash_total
    elif main_sheet['H12'].value == None:
        main_sheet['H12'] = cash_total

    # UBER EATS
    uber_eats = xls.search_value_in_column(ws, 'Uber Eats', 'B')
    x = uber_eats[1]
    y = payment_summary_total[0]
    z = y + str(x)
    uber_eats_total = ws[z].value
    # Loops through each of the files in order fill in the data
    if main_sheet['B13'].value == None:
        main_sheet['B13'] = uber_eats_total
    elif main_sheet['C13'].value == None:
        main_sheet['C13'] = uber_eats_total
    elif main_sheet['D13'].value == None:
        main_sheet['D13'] = uber_eats_total
    elif main_sheet['E13'].value == None:
        main_sheet['E13'] = uber_eats_total
    elif main_sheet['F13'].value == None:
        main_sheet['F13'] = uber_eats_total
    elif main_sheet['G13'].value == None:
        main_sheet['G13'] = uber_eats_total
    elif main_sheet['H13'].value == None:
        main_sheet['H13'] = uber_eats_total

    # TIPS PAID
    tips_column = xls.search_value_in_row(ws, 'Tips', payment_summary[1])
    gratuity_column = xls.search_value_in_row(ws, 'Gratuity', payment_summary[1])
    other_payment = xls.search_value_in_column(ws, 'Other', 'B')
    tips_coords = tips_column[0] + str(other_payment[1] + 1)
    gratuity_coords = gratuity_column[0] + str(other_payment[1] + 1)
    tips = ws[tips_coords].value
    gratuity = ws[gratuity_coords].value
    tips_paid = tips + gratuity
    # Loops through each of the files in order fill in the data
    if main_sheet['B14'].value == None:
        main_sheet['B14'] = tips_paid
    elif main_sheet['C14'].value == None:
        main_sheet['C14'] = tips_paid
    elif main_sheet['D14'].value == None:
        main_sheet['D14'] = tips_paid
    elif main_sheet['E14'].value == None:
        main_sheet['E14'] = tips_paid
    elif main_sheet['F14'].value == None:
        main_sheet['F14'] = tips_paid
    elif main_sheet['G14'].value == None:
        main_sheet['G14'] = tips_paid
    elif main_sheet['H14'].value == None:
        main_sheet['H14'] = tips_paid

    # S-FOOD
    sales_categories = xls.search_value_in_column(ws, 'Sales Categories')
    gross_amnt = xls.search_value_in_row(ws, 'Gross Amt', sales_categories[1])
    food = xls.search_value_in_column(ws, 'Food', 'B')
    temp = gross_amnt[0] + str(food[1])
    temp2 = gross_amnt[0] + str(food[1]- 1)
    temp_total = ws[temp].value
    temp_total2 = ws[temp2].value
    final = temp_total + temp_total2
    # Loops through each of the files in order fill in the data
    if main_sheet['B15'].value == None:
        main_sheet['B15'] = final
    elif main_sheet['C15'].value == None:
        main_sheet['C15'] = final
    elif main_sheet['D15'].value == None:
        main_sheet['D15'] = final
    elif main_sheet['E15'].value == None:
        main_sheet['E15'] = final
    elif main_sheet['F15'].value == None:
        main_sheet['F15'] = final
    elif main_sheet['G15'].value == None:
        main_sheet['G15'] = final
    elif main_sheet['H15'].value == None:
        main_sheet['H15'] = final

    #S-BAR
    no_category = xls.search_value_in_column(ws, 'No Category', 'B')
    temp =  gross_amnt[0] + str(no_category[1] + 1)
    temp3 = ws[temp].value
    sub = temp3 - final
    # Loops through each of the files in order fill in the data
    if main_sheet['B16'].value == None:
        main_sheet['B16'] = sub
    elif main_sheet['C16'].value == None:
        main_sheet['C16'] = sub
    elif main_sheet['D16'].value == None:
        main_sheet['D16'] = sub
    elif main_sheet['E16'].value == None:
        main_sheet['E16'] = sub
    elif main_sheet['F16'].value == None:
        main_sheet['F16'] = sub
    elif main_sheet['G16'].value == None:
        main_sheet['G16'] = sub
    elif main_sheet['H16'].value == None:
        main_sheet['H16'] = sub

    # COMP PROMO
    menu_item_discounts = xls.search_value_in_column(ws, 'Menu Item Discounts')
    comp_item = xls.search_value_in_column(ws, 'Comp % Item', 'B')
    menu_amount = xls.search_value_in_row(ws, 'Amount', menu_item_discounts[1])
    comp_item_coords = menu_amount[0] + str(comp_item[1])
    early_bird_coords = menu_amount[0] + str(comp_item[1] + 1)
    happy_hr_coords = menu_amount[0] + str(comp_item[1] + 2)
    comp_item = ws[comp_item_coords].value
    early_bird = ws[early_bird_coords].value
    happy_hr = ws[happy_hr_coords].value
    final = comp_item + early_bird + happy_hr
    # Loops through each of the files in order fill in the data
    if main_sheet['B18'].value == None:
        main_sheet['B18'] = final
    elif main_sheet['C18'].value == None:
        main_sheet['C18'] = final
    elif main_sheet['D18'].value == None:
        main_sheet['D18'] = final
    elif main_sheet['E18'].value == None:
        main_sheet['E18'] = final
    elif main_sheet['F18'].value == None:
        main_sheet['F18'] = final
    elif main_sheet['G18'].value == None:
        main_sheet['G18'] = final
    elif main_sheet['H18'].value == None:
        main_sheet['H18'] = final

    # MGR DISCOUNT
    check_discounts = xls.search_value_in_column(ws, 'Check Discounts')
    manager_comp = xls.search_value_in_column(ws, 'Manager Comp - Check', 'B')
    price = menu_amount[0] + str(manager_comp[1])
    final = ws[price].value
    # Loops through each of the files in order fill in the data
    if main_sheet['B19'].value == None:
        main_sheet['B19'] = final
    elif main_sheet['C19'].value == None:
        main_sheet['C19'] = final
    elif main_sheet['D19'].value == None:
        main_sheet['D19'] = final
    elif main_sheet['E19'].value == None:
        main_sheet['E19'] = final
    elif main_sheet['F19'].value == None:
        main_sheet['F19'] = final
    elif main_sheet['G19'].value == None:
        main_sheet['G19'] = final
    elif main_sheet['H19'].value == None:
        main_sheet['H19'] = final

    #SERV.DISCOUNT
    price = menu_amount[0] + str(manager_comp[1] - 1)
    final = ws[price].value
    # Loops through each of the files in order fill in the data
    if main_sheet['B21'].value == None:
        main_sheet['B21'] = final
    elif main_sheet['C21'].value == None:
        main_sheet['C21'] = final
    elif main_sheet['D21'].value == None:
        main_sheet['D21'] = final
    elif main_sheet['E21'].value == None:
        main_sheet['E21'] = final
    elif main_sheet['F21'].value == None:
        main_sheet['F21'] = final
    elif main_sheet['G21'].value == None:
        main_sheet['G21'] = final
    elif main_sheet['H21'].value == None:
        main_sheet['H21'] = final

    # QSA
    comp_item = xls.search_value_in_column(ws, 'Comp % Item', 'B')
    qsa_coords = menu_amount[0] + str(comp_item[1] + 3)
    qsa = ws[qsa_coords].value
    # Loops through each of the files in order fill in the data
    if main_sheet['B22'].value == None:
        main_sheet['B22'] = qsa
    elif main_sheet['C22'].value == None:
        main_sheet['C22'] = qsa
    elif main_sheet['D22'].value == None:
        main_sheet['D22'] = qsa
    elif main_sheet['E22'].value == None:
        main_sheet['E22'] = qsa
    elif main_sheet['F22'].value == None:
        main_sheet['F22'] = qsa
    elif main_sheet['G22'].value == None:
        main_sheet['G22'] = qsa
    elif main_sheet['H22'].value == None:
        main_sheet['H22'] = qsa

    # SALES TAX
    taxes = xls.search_value_in_column(ws, 'Tax', 'B')
    price = menu_amount[0] + str(taxes[1] + 1)
    final = ws[price].value
    # Loops through each of the files in order fill in the data
    if main_sheet['B23'].value == None:
        main_sheet['B23'] = final
    elif main_sheet['C23'].value == None:
        main_sheet['C23'] = final
    elif main_sheet['D23'].value == None:
        main_sheet['D23'] = final
    elif main_sheet['E23'].value == None:
        main_sheet['E23'] = final
    elif main_sheet['F23'].value == None:
        main_sheet['F23'] = final
    elif main_sheet['G23'].value == None:
        main_sheet['G23'] = final
    elif main_sheet['H23'].value == None:
        main_sheet['H23'] = final
    
    # Saves the New Weekly Sales Report
    main_workbook.save("Weekly_Sales.xlsx")
