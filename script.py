import os
from pathlib import Path

import win32com.client as client
from openpyxl import Workbook, load_workbook
from openpyxl.styles import (Alignment, Border, Font, PatternFill, Protection,
                             Side)

import models.xls_to_xlsx as xls

# Clears the new versions folder 
""" Change Path to your path to new versions folder!!! """
[f.unlink() for f in Path("C:\\Users\\12392\\Desktop\\Github\\Toast-Expense-Weekly_Sales\\new version").glob("*") if f.is_file()]

# Converts the xls file to xlsx format
xls.converter()

# Creates the new Workbook
main_workbook = load_workbook(os.getcwd() + "\\template\\template.xlsx")
main_sheet = main_workbook.active
main_sheet.title = "Summary" # Fill in the date with the last date

dir = os.getcwd() + '/new version/'

filess = os.listdir(dir)
file_list = [files for files in filess if '~' not in files[0]]
for files in file_list:
    if 'False' in files:
        continue
    if files[0] != '~':
        wb = load_workbook(dir + files)
        # Opens the first sheet in the workbook
        ws = wb['Summary']
        # Gets the Date
        if files == file_list[0]:
            temp_date = ws['A2'].value
            split = temp_date.split('-')
            date = split[0].strip()
            # Checks where to put date
            main_sheet['A4'] = date

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
        if files == file_list[0]:
            main_sheet['B7'] = total_deposit
        elif files == file_list[1]:
            main_sheet['C7'] = total_deposit
        elif files == file_list[2]:
            main_sheet['D7'] = total_deposit
        elif files == file_list[3]:
            main_sheet['E7'] = total_deposit
        elif files == file_list[4]:
            main_sheet['F7'] = total_deposit
        elif files == file_list[5]:
            main_sheet['G7'] = total_deposit
        elif files == file_list[6]:
            main_sheet['H7'] = total_deposit

        # Gift Cards SOlD 
        sales_summary = xls.search_value_in_column(ws, 'Sales Summary')
        deferred = xls.search_value_in_row(ws, 'Deferred', sales_summary[1])
        x = int(sales_summary[1]) + 1
        y = deferred[0]
        z = y + str(x)
        if 'INone' != z:
            gift_sold_value = ws[z].value
            gift_sold = float(gift_sold_value.strip('$'))
            # Loops through each of the files in order fill in the data
            if gift_sold > 0:
                if files == file_list[0]:
                    main_sheet['B8'] = gift_sold
                elif files == file_list[1]:
                    main_sheet['C8'] = gift_sold
                elif files == file_list[2]:
                    main_sheet['D8'] = gift_sold
                elif files == file_list[3]:
                    main_sheet['E8'] = gift_sold
                elif files == file_list[4]:
                    main_sheet['F8'] = gift_sold
                elif files == file_list[5]:
                    main_sheet['G8'] = gift_sold
                elif files == file_list[6]:
                    main_sheet['H8'] = gift_sold

    # Gift Cards REDEEMED
    gift_card_coords = xls.search_value_in_column(ws, 'Gift Card', 'B')
    x = gift_card_coords[1]
    if 'None' not in str(x):
        y = payment_summary_total[0]
        z = y + str(x)
        gift_card_total = ws[z].value
        # Loops through each of the files in order fill in the data
        if files == file_list[0]:
            main_sheet['B9'] = gift_card_total
        elif files == file_list[1]:
            main_sheet['C9'] = gift_card_total
        elif files == file_list[2]:
            main_sheet['D9'] = gift_card_total
        elif files == file_list[3]:
            main_sheet['E9'] = gift_card_total
        elif files == file_list[4]:
            main_sheet['F9'] = gift_card_total
        elif files == file_list[5]:
            main_sheet['G9'] = gift_card_total
        elif files == file_list[6]:
            main_sheet['H9'] = gift_card_total

    # CASH SALES
    cash_location = xls.search_value_in_column(ws, 'Cash', 'B')
    x = cash_location[1]
    y = payment_summary_total[0]
    z = y + str(x)
    cash_total = ws[z].value
    print(cash_total)
    # Loops through each of the files in order fill in the data
    if cash_total > 0:
        if files == file_list[0]:
            main_sheet['B12'] = cash_total
        elif files == file_list[1]:
            main_sheet['C12'] = cash_total
        elif files == file_list[2]:
            main_sheet['D12'] = cash_total
        elif files == file_list[3]:
            main_sheet['E12'] = cash_total
        elif files == file_list[4]:
            main_sheet['F12'] = cash_total
        elif files == file_list[5]:
            main_sheet['G12'] = cash_total
        elif files == file_list[6]:
            main_sheet['H12'] = cash_total

    # UBER EATS
    uber_eats = xls.search_value_in_column(ws, 'Uber Eats', 'B')
    x = uber_eats[1]
    if 'None' not in str(x):
        y = payment_summary_total[0]
        z = y + str(x)
        uber_eats_total = ws[z].value
        # Loops through each of the files in order fill in the data
        if uber_eats_total > 0:
            if files == file_list[0]:
                main_sheet['B13'] = uber_eats_total
            elif files == file_list[1]:
                main_sheet['C13'] = uber_eats_total
            elif files == file_list[2]:
                main_sheet['D13'] = uber_eats_total
            elif files == file_list[3]:
                main_sheet['E13'] = uber_eats_total
            elif files == file_list[4]:
                main_sheet['F13'] = uber_eats_total
            elif files == file_list[5]:
                main_sheet['G13'] = uber_eats_total
            elif files == file_list[6]:
                main_sheet['H13'] = uber_eats_total

    # TIPS PAID
    tips_coords = 'E5'
    gratuity_coords = 'D5'
    tips_value = ws[tips_coords].value
    gratuity_value = ws[gratuity_coords].value
    tips = tips_value[1:]
    gratuity = gratuity_value[1:]
    if ',' in tips:
        tipped = tips.replace(',','')
        if ',' in gratuity:
            gratuitys = gratuity.replace(',','')
            tips_paid = round(float(tipped) + float(gratuitys), 3)
        else:
            tips_paid = round(float(tipped) + float(gratuity),3)
    else:
        tips_paid = round(float(tips) + float(gratuity),3)
    # Loops through each of the files in order fill in the data
    if tips_paid > 0:
        if files == file_list[0]:
            main_sheet['B14'] = tips_paid
        elif files == file_list[1]:
            main_sheet['C14'] = tips_paid
        elif files == file_list[2]:
            main_sheet['D14'] = tips_paid
        elif files == file_list[3]:
            main_sheet['E14'] = tips_paid
        elif files == file_list[4]:
            main_sheet['F14'] = tips_paid
        elif files == file_list[5]:
            main_sheet['G14'] = tips_paid
        elif files == file_list[6]:
            main_sheet['H14'] = tips_paid

    # S-FOOD
    sales_categories = xls.search_value_in_column(ws, 'Sales Categories')
    food = xls.search_value_in_column(ws, 'Food', 'B')
    food_amnt_coords = 'E' + str(food[1])
    na_beverage_amnt_coords = 'E' + str(food[1]- 1)
    food_total = ws[food_amnt_coords].value
    na_beverage_total = ws[na_beverage_amnt_coords].value
    final = food_total + na_beverage_total
    # Loops through each of the files in order fill in the data
    if final > 0:
        if files == file_list[0]:
            main_sheet['B15'] = final
        elif files == file_list[1]:
            main_sheet['C15'] = final
        elif files == file_list[2]:
            main_sheet['D15'] = final
        elif files == file_list[3]:
            main_sheet['E15'] = final
        elif files == file_list[4]:
            main_sheet['F15'] = final
        elif files == file_list[5]:
            main_sheet['G15'] = final
        elif files == file_list[6]:
            main_sheet['H15'] = final

    #S-BAR
    no_category = xls.search_value_in_column(ws, 'No Category', 'B')
    food_beverage_total = 'E' + str(no_category[1] + 1)
    temp3 = ws[food_beverage_total].value
    sub = temp3 - final
    # Loops through each of the files in order fill in the data
    if sub > 0:
        if files == file_list[0]:
            main_sheet['B16'] = sub
        elif files == file_list[1]:
            main_sheet['C16'] = sub
        elif files == file_list[2]:
            main_sheet['D16'] = sub
        elif files == file_list[3]:
            main_sheet['E16'] = sub
        elif files == file_list[4]:
            main_sheet['F16'] = sub
        elif files == file_list[5]:
            main_sheet['G16'] = sub
        elif files == file_list[6]:
            main_sheet['H16'] = sub

    # COMP PROMO
    comp_item_location = xls.search_value_in_column(ws, 'Comp % Item', 'B')
    early_bird_location = xls.search_value_in_column(ws, 'EARLY BIRD', 'B')
    happy_hr_location = xls.search_value_in_column(ws, 'HAPPY HOUR', 'B')
    open_check_dollar_location = xls.search_value_in_column(ws, 'Open $ Check', 'B')
    checks = 0
    # Finds the location using the row
    if None not in comp_item_location:
        comp_item_coords = 'D' + str(comp_item_location[1])
        checks += ws[comp_item_coords].value
    if None not in early_bird_location:
        early_bird_coords = 'D' + str(early_bird_location[1])
        checks += ws[early_bird_coords].value
    if None not in happy_hr_location:
        happy_hr_coords = 'D' + str(happy_hr_location[1])
        checks += ws[happy_hr_coords].value
    if None not in open_check_dollar_location:
        open_check_coords = 'D' + str(open_check_dollar_location[1])
        checks += ws[open_check_coords].value
    # Loops through each of the files in order fill in the data
    if files == file_list[0]:
        main_sheet['B18'] = checks
    elif files == file_list[1]:
        main_sheet['C18'] = checks
    elif files == file_list[2]:
        main_sheet['D18'] = checks
    elif files == file_list[3]:
        main_sheet['E18'] = checks
    elif files == file_list[4]:
        main_sheet['F18'] = checks
    elif files == file_list[5]:
        main_sheet['G18'] = checks
    elif files == file_list[6]:
        main_sheet['H18'] = checks

    # MGR DISCOUNT
    check_discounts = xls.search_value_in_column(ws, 'Check Discounts')
    manager_comp = xls.search_value_in_column(ws, 'Manager Comp - Check', 'B')
    if None not in manager_comp:
        price = 'D' + str(manager_comp[1])
        final = ws[price].value
        # Loops through each of the files in order fill in the data
        if files == file_list[0]:
            main_sheet['B19'] = final
        elif files == file_list[1]:
            main_sheet['C19'] = final
        elif files == file_list[2]:
            main_sheet['D19'] = final
        elif files == file_list[3]:
            main_sheet['E19'] = final
        elif files == file_list[4]:
            main_sheet['F19'] = final
        elif files == file_list[5]:
            main_sheet['G19'] = final
        elif files == file_list[6]:
            main_sheet['H19'] = final

    #SERV.DISCOUNT
    employee_discount = xls.search_value_in_column(ws, 'Employee Discount - Check', 'B')
    if None not in employee_discount:
        price = 'D' + str(employee_discount[1])
        final = ws[price].value
        # Loops through each of the files in order fill in the data
        if files == file_list[0]:
            main_sheet['B21'] = final
        elif files == file_list[1]:
            main_sheet['C21'] = final
        elif files == file_list[2]:
            main_sheet['D21'] = final
        elif files == file_list[3]:
            main_sheet['E21'] = final
        elif files == file_list[4]:
            main_sheet['F21'] = final
        elif files == file_list[5]:
            main_sheet['G21'] = final
        elif files == file_list[6]:
            main_sheet['H21'] = final

    # QSA
    qsa_location = xls.search_value_in_column(ws, 'QSA', 'B')
    if None not in qsa_location:
        qsa_coords = 'D' + str(qsa_location[1])
        qsa = ws[qsa_coords].value
        # Loops through each of the files in order fill in the data
        if files == file_list[0]:
            main_sheet['B22'] = qsa
        elif files == file_list[1]:
            main_sheet['C22'] = qsa
        elif files == file_list[2]:
            main_sheet['D22'] = qsa
        elif files == file_list[3]:
            main_sheet['E22'] = qsa
        elif files == file_list[4]:
            main_sheet['F22'] = qsa
        elif files == file_list[5]:
            main_sheet['G22'] = qsa
        elif files == file_list[6]:
            main_sheet['H22'] = qsa

    # SALES TAX
    price = ws['C5'].value
    takes = price[1:]
    final = float(takes)
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

# Clears the new versions folder
[f.unlink() for f in Path("C:\\Users\\12392\\Desktop\\Github\\Toast-Expense-Weekly_Sales\\new version").glob("*") if f.is_file()]

# Saves the New Weekly Sales Report
main_workbook.save("Weekly_Sales.xlsx")
