#! python 2.7
# -*- coding: cp1252 -*-
# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils import column_index_from_string

# variables de entrada
wb = openpyxl.load_workbook('contabilidad.xlsx')
sheet = wb['hoja_trabajo_Contab']

auxiliares = ['6051','6052','6053','6101','6103','6401','6408','6409','6410','6411','6412','6413','6414','6415','6416','6418','6451','6452','6501','6502','6503','6551','6553','6554','6555','6556','6557','6558','6559','6601','6602','6603','6604','6605','6606','6607','6651','6652','6701','6702','6704','6705','6751','6753','6754','6801','6802','6804','7152','7157','9152','9601','9603','9604','9605','9606','9607','9701','9702','65590']


requerimiento = {'4001': '1-NIVEL CENTRAL',  #------------------------------------------------->   modificar con HByS
                 '4010': '2-FONDOS ESPECIALES',
                 '4011': '2-FONDOS ESPECIALES',
                 '4012': '2-FONDOS ESPECIALES',
                 '4013': '2-FONDOS ESPECIALES',
                 '4014': '2-FONDOS ESPECIALES',
                 '4002': '6-UNISALUD',
                 '4003': 'NA',
                 '4011': 'NA',
                 '4012': 'NA',
                 '1013': 'NA',
                 '4014': 'NA',
                 '1060': '3-UGIS',
                 '4061': 'NA',
                 '4062': 'NA',
                 '4063': 'NA',
                 '4064': 'NA',
                 '4065': 'NA',
                 '1005': '4-UNIMEDIOS',
                 '1011': '4-UNIMEDIOS',
                 '1004': '5-UNIBIBLOS',
                 }

reqValidos = ['4001','4010','4011','4012','4013','4014','4002']
reqFondosEspeciales = ['4010','4011','4012','4013','4014']


def recorrer():
 con = 0

 for row in range(3, sheet.max_row + 1):
  if( str(sheet['B' + str(row)].value) in auxiliares and str(sheet['A' + str(row)].value) in reqValidos ):
   #print 'entra1'
   fila = {
       "Empresa": str(sheet['A' + str(row)].value), 
       "NITaux" : str(sheet['B' + str(row)].value),
       "Cuenta" : str(sheet['C' + str(row)].value)
   }
   #con el comprobante 18
   if (fila["Empresa"] in reqFondosEspeciales):
    if (sheet['H' + str(row)].value is not None):
     for row2 in range(0, 5):
      if (round(float(sheet['H' + str(row)].value), 0) <> round(float(sumatoriaD(fila, get_column_letter(column_index_from_string('T') + row2), 18)), 0) and float(sheet['H' + str(row)].value) <> 0):  # == para las que si coinciden
       con += 1
       print ' * ' + fila["Empresa"] + " " + fila["NITaux"] + " " + str(sheet['C' + str(row)].value) + " Debito (Hoja ByS): " + str(sumatoriaD(fila, get_column_letter(column_index_from_string('T') + row2) , 18)) + " - Valor en hoja contable : " + str(sheet['H' + str(row)].value)

      if (round(float(sheet['I' + str(row)].value), 0) <> round(float(sumatoriaC(fila, get_column_letter(column_index_from_string('T') + row2), 18)), 0) and float(sheet['I' + str(row)].value) <> 0):
       con += 1
       print ' * ' + fila["Empresa"] + " " + fila["NITaux"] + " " + str(sheet['C' + str(row)].value) + " Credito (Hoja ByS): " + str(sumatoriaC(fila, get_column_letter(column_index_from_string('T') + row2), 18)) + " - Valor en hoja contable : " + str(sheet['I' + str(row)].value)

   else:
    if (sheet['H' + str(row)].value is not None):
     if (round(float(sheet['H' + str(row)].value), 0) <> round(float(sumatoriaD(fila, 'T', 18)), 0) and float(sheet['H' + str(row)].value) <> 0):  # == para las que si coinciden
      con += 1
      print ' * ' + fila["Empresa"] + " " + fila["NITaux"] + " " + str(sheet['C' + str(row)].value) + " Debito (Hoja ByS): " + str(sumatoriaD(fila, 'T', 18)) + " - Valor en hoja contable : " + str(sheet['H' + str(row)].value)

     if (round(float(sheet['I' + str(row)].value), 0) <> round(float(sumatoriaC(fila, 'T', 18)), 0) and float(sheet['I' + str(row)].value) <> 0):
      con += 1
      print ' * ' + fila["Empresa"] + " " + fila["NITaux"] + " " + str(sheet['C' + str(row)].value) + " Credito (Hoja ByS): " + str(sumatoriaC(fila, 'T', 18)) + " - Valor en hoja contable : " + str(sheet['I' + str(row)].value)

   #con el comprobante 16
   if (fila["Empresa"] in reqFondosEspeciales):
    if (sheet['F' + str(row)].value is not None):
     for row2 in range(0, 5):
      if (round(float(sheet['F' + str(row)].value), 0) <> round(float(sumatoriaD(fila, get_column_letter(column_index_from_string('T') + row2) , 16)), 0) and float(sheet['F' + str(row)].value) <> 0):  # == para las que si coinciden
       con += 1
       print ' * ' + fila["Empresa"] + " " + fila["NITaux"] + " " + str(sheet['C' + str(row)].value) + " Debito (Hoja ByS): " + str(sumatoriaD(fila, get_column_letter(column_index_from_string('T') + row2) , 16)) + " - Valor en hoja contable : " + str(sheet['F' + str(row)].value)

      if (round(float(sheet['G' + str(row)].value), 0) <> round(float(sumatoriaC(fila, get_column_letter(column_index_from_string('T') + row2), 16)), 0) and float(sheet['G' + str(row)].value) <> 0):
       con += 1
       print ' * ' + fila["Empresa"] + " " + fila["NITaux"] + " " + str(sheet['C' + str(row)].value) + " Credito (Hoja ByS): " + str(sumatoriaC(fila, get_column_letter(column_index_from_string('T') + row2), 16)) + " - Valor en hoja contable : " + str(sheet['G' + str(row)].value)

   else:
    if (sheet['F' + str(row)].value is not None):
     if (round(float(sheet['F' + str(row)].value), 0) <> round(float(sumatoriaD(fila, 'T', 16)), 0) and float(sheet['F' + str(row)].value) <> 0):  # == para las que si coinciden
      con += 1
      print ' * ' + fila["Empresa"] + " " + fila["NITaux"] + " " + str(sheet['C' + str(row)].value) + " Debito (Hoja ByS): " + str(sumatoriaD(fila, 'T', 16)) + " - Valor en hoja contable : " + str(sheet['F' + str(row)].value)

     if (round(float(sheet['G' + str(row)].value), 0) == round(float(sumatoriaC(fila, 'T', 16)), 0) and float(sheet['G' + str(row)].value) <> 0):
      con += 1
      print ' * ' + fila["Empresa"] + " " + fila["NITaux"] + " " + str(sheet['C' + str(row)].value) + " Credito (Hoja ByS): " + str(sumatoriaC(fila, 'T', 16)) + " - Valor en hoja contable : " + str(sheet['G' + str(row)].value)

 if (con == 0):
  print " *** VALIDACION EXITOSA *** "
 else:
  print str(con) + " resultados erroneos"


def sumatoriaD(fila, col, comprobante):
    sheet = wb['ByS']
    sumatoria = 0

    for row in range(2, sheet.max_row + 1):
     if( sheet['O' + str(row)].value is not None and str(sheet['N' + str(row)].value) is not None):
      if (str(sheet['F' + str(row)].value) == requerimiento[fila["Empresa"]] and int(sheet['O' + str(row)].value) == int(fila["NITaux"]) and str(sheet['N' + str(row)].value) == fila["Cuenta"]):
       if (str(sheet[col + str(row)].value) <> '0' and sheet[col + str(row)].value is not None and sheet['L' + str(row)].value == comprobante ):
        sumatoria += float(sheet[col + str(row)].value)

    return sumatoria


def sumatoriaC(fila, col, comprobante):
    sheet = wb['ByS']
    sumatoria = 0

    for row in range(2, sheet.max_row + 1):
     if (sheet['O' + str(row)].value is not None and str(sheet['N' + str(row)].value) is not None):
      if (str(sheet['F' + str(row)].value) == requerimiento[fila["Empresa"]] and int( sheet['R' + str(row)].value) == int(fila["NITaux"]) and str(sheet['Q' + str(row)].value) == fila["Cuenta"]):
       if (str(sheet[col + str(row)].value) <> '0' and sheet[col + str(row)].value is not None and sheet['L' + str(row)].value == comprobante):
        sumatoria += float(sheet[col + str(row)].value)

    return sumatoria


recorrer()
