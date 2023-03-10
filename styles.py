# Styles module for formatting cells in output Excel file

from openpyxl.styles import Font, Alignment, numbers


def left_aligned_cell(cell):
    cell.font = Font(name='Calibri', size=14)
    cell.alignment = Alignment(horizontal='left', vertical='center')


def centered_cell(cell):
    cell.font = Font(name='Calibri', size=14)
    cell.alignment = Alignment(horizontal='center', vertical='center')


def currency_cell(cell):
    cell.font = Font(name='Calibri', size=14)
    cell.alignment = Alignment(horizontal='right', vertical='center')
    cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE