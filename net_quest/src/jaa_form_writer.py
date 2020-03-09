import xlsxwriter
from pathlib import Path
import os

def write_jaa_intake_form(global_items):

    prospect = global_items[0][1]

    file_name = "_".join(['CH.JAAForm' , prospect ,'[DATE]' , '.xlsx' ])

    output_directory = Path(os.environ['HOME'] + '/Desktop/net_quest_output/intake_forms/' + file_name)

    workbook = xlsxwriter.Workbook(output_directory)

    worksheet = workbook.add_worksheet()

    # Creates four cell formats
    cell_format1 = workbook.add_format({
        'bold': '1',
        'font_size': '10',
        'font_name': 'Arial'})

    cell_format2 = workbook.add_format({
        'font_size': '10',
        'font_name': 'Arial'})

    cell_format3 = workbook.add_format({
        'bold': '1',
        'font_size': '10',
        'font_name': 'Arial',
        'font_color': 'Blue'})

    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#119922',
        'font_name': 'Arial'})

    # Merges and formats header
    worksheet.set_column('A:B', 44)
    worksheet.merge_range('A1:B1', 'Collective Health', merge_format)

    # Write first column (fields) for JAA intake form
    worksheet.write(4, 0, 'Employer Name', cell_format1)
    worksheet.write(5, 0, 'Employer Location / Headquarter', cell_format1)
    worksheet.write(6, 0, 'Business (SIC if available)', cell_format1)
    worksheet.write(7, 0, 'Number of Employees', cell_format1)
    worksheet.write(8, 0, 'Number of Covered Employees', cell_format1)
    worksheet.write(9, 0, 'Number of total Members (employees and dependents)', cell_format1)
    worksheet.write(10, 0, 'Census', cell_format1)
    worksheet.write(11, 0, 'Current PPO/HMO Network', cell_format1)
    worksheet.write(12, 0, 'Current Administrator (Carrier & TPA)', cell_format1)
    worksheet.write(13, 0, 'Proposed Network (CA Only or JAA *)', cell_format1)
    worksheet.write(14, 0, 'Number of out-of-state Employees (If requesting JAA)', cell_format1)
    worksheet.write(15, 0, 'Proposed Effective Date', cell_format1)
    worksheet.write(16, 0, 'Broker or Consultant Name', cell_format1)
    worksheet.write(17, 0, 'Broker or Consultant Location', cell_format1)

    # Writes JAA callout
    worksheet.write(20, 0, '* JAA includes BlueCard access', cell_format3)

    # Write second column (values) for JAA intake form using global variables from render_text.py
    worksheet.write(4, 1, str(global_items[0][1]), cell_format2) # prospect
    worksheet.write(5, 1, str(global_items[18][1]), cell_format2)
    worksheet.write(6, 1, ' ', cell_format2)
    worksheet.write(7, 1, str(global_items[13][1]), cell_format2) # ees
    worksheet.write(8, 1, str(global_items[14][1]), cell_format2) # covered es

    worksheet.write(9, 1, 'See Census', cell_format2) # total members
    worksheet.write(10, 1, 'Attached', cell_format2) # see census

    if global_items[17][1] == True: # Kaiser
        worksheet.write(11, 1, 'Kaiser', cell_format2)
    else: worksheet.write(11, 1, '', cell_format2)

    worksheet.write(12, 1, str(global_items[15][1]), cell_format2) # current medical carrier
    worksheet.write(13, 1, 'JAA', cell_format2) # proposed network
    worksheet.write(14, 1, 'See Census', cell_format2) # out of state employees
    worksheet.write(15, 1, str(global_items[6][1]), cell_format2) # start date
    worksheet.write(16, 1, str(global_items[10][1]), cell_format2) # broker contact
    worksheet.write(17, 1, str(global_items[9][1]), cell_format2) # broker address

    workbook.close()

