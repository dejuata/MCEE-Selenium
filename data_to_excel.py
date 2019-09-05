# -*- coding: utf-8 -*-
import validaciones
import openpyxl
import os

from selenium import webdriver


# Read file excel
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

  # count = 0

  for r in range(2, rows+1):
    data_cell = sheet.cell(row=r, column=column).value    
    fecha = validaciones.generar_fecha(data_cell)    
    sheet.cell(row=r, column=column).value = fecha

  # Cerrar archivo de texto y excel
  workbook.save(path)
  # print('Faltan: {}'.format(count))
  print("data success")

completar_comuna(workbook, sheet, 12, 11088)

def validar_numero(workbook, sheet, column, rows, log):
  
  file = open(log, "w")

  for r in range(2, rows+1):
    print(r)
    data_cell = sheet.cell(row=r, column=column).value
    if not (data_cell == None or data_cell == " "):
      if not validaciones.validar_correo(data_cell.lower()):
        file.write("Row: {} failed".format(r) + os.linesep)
    
 # Cerrar archivo de texto y excel
  file.close()
  print("data success")

# validar_numero(workbook, sheet, 27, 11088, './log.txt')