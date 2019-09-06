# -*- coding: utf-8 -*-
import validaciones
import openpyxl
import os

from selenium import webdriver


# Read file excel
# path = './BD2.xlsx'
path = './CB1.xlsx'
SHEET_NAME = 'TOTAL BENEFICIARIOS EN IEO'

# Cargar Archivo
workbook = openpyxl.load_workbook(path)
sheet = workbook.get_sheet_by_name(SHEET_NAME)

# Obtener filas del archivo
rows = sheet.max_row
# Obtener columnas del archivo
# cols = sheet.max_column

def completar_comuna(workbook, sheet, column, rows):

  count = 0

  for r in range(2, rows+1):
    data_cell = sheet.cell(row=r, column=column).value
    
    # if not data_cell == None:
    #   if data_cell.lower() == "Vereda Loma Alta".lower():
    #     sheet.cell(row=r, column=column+1).value = "No vive en Cali"
    #     sheet.cell(row=r, column=column+2).value = "Vereda Loma Alta"

    comuna = sheet.cell(row=r, column=column+1).value
    if (comuna == None or comuna == " "):
      count += 1

  # Cerrar archivo de texto y excel
  workbook.save(path)
  print('Faltan: {}'.format(count))
  print("data success")

# completar_comuna(workbook, sheet, 23, 6308)

def validar_numero(workbook, sheet, column, rows, log):
  
  file = open(log, "w")

  for r in range(2, rows+1):
    print(r)
    data_cell = sheet.cell(row=r, column=column).value
    if not (data_cell == None or data_cell == " "):
      if not validaciones.validar_lgtbi(data_cell):
        # sheet.cell(row=r, column=column).value = '2'
        file.write("Row: {} failed".format(r) + os.linesep)
    
 # Cerrar archivo de texto y excel
  # workbook.save(path)
  file.close()
  print("data success")

validar_numero(workbook, sheet, 15, 12002, './log.txt')