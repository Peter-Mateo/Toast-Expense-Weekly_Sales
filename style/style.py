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

