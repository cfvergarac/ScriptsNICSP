#! python 2.7
# -*- coding: cp1252 -*-
# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils import column_index_from_string


wb = openpyxl.load_workbook('Plantilla contable consolidada (solo valores) 28022018.xlsx')
sheet = wb['Hoja1']
contador = 0

ajustes = {'1': 'AN',
           '2': 'AW',
           '3': 'BG',
           '4': 'BP',
           '5': 'BY',
           '6': 'AW',
           '7': 'BG',
           '8': 'BP',
           '9': 'BY',
           '10': 'AN',
           '11': 'AN',
           '14': 'AW',
           '15': 'AW'}


requerimiento = {'1-NIVEL CENTRAL': '4001',
                 '2-FONDOS ESPECIALES': ['4010','4011','4012','4013','4014'],
                 '6-UNISALUD': '4002'}

requerimientoInvalido =  ['3-UGIS','4-UNIMEDIOS','5-UNIBIBLOS']


fila = {
            "numajuste": str(sheet['A13'].value),
            "grupo": str(sheet['O13' ].value)[0:3],
            "subgrupo": str((sheet['E13'].value).split("-")[0]),
            "requerimiento": str(requerimiento[sheet['F13'].value]),
            "procedencia": str((sheet['G13' ].value).split("-")[0]),
            "categoria": str((sheet['I13'].value).split("-")[0]),
            "celda": str(ajustes[str(sheet['A13'].value)]),
            "estado": str((sheet['H13'].value).split("-")[0]),
            "Partida": str(sheet['N13'].value),
            "auxPartida": str(sheet['O13'].value),
            "ContraPartida": str(sheet['Q13'].value),
            "auxContraPartida": str(sheet['R13'].value)
}
#sheet.max_row + 1

def recorrer():
 for row in range(2, sheet.max_row + 1):
  if (sheet['N' + str(row)].value is not None and sheet['O' + str(row)].value is not None and sheet['Q' + str(row)].value is not None and sheet['R' + str(row)].value is not None and sheet['F' + str(row)].value not in requerimientoInvalido):
   if (str(sheet['F' + str(row)].value) == '2-FONDOS ESPECIALES'):

    for row2 in range(0, 5):
     fila = {
               "numajuste": str(sheet['A' + str(row)].value),
               "grupo": str(sheet['O' + str(row)].value)[0:3],
               "subgrupo": str((sheet['E' + str(row)].value).split("-")[0]),
               "requerimiento": str(requerimiento['2-FONDOS ESPECIALES'][row2]),
               "procedencia": str((sheet['G' + str(row)].value).split("-")[0]),
               "categoria": str((sheet['I' + str(row)].value).split("-")[0]),
               "celda": str(ajustes[str(sheet['A' + str(row)].value)]),
               "estado": str((sheet['H' + str(row)].value).split("-")[0]),
               "Partida": str(sheet['N' + str(row)].value),
               "auxPartida": str(sheet['O' + str(row)].value),
               "ContraPartida": str(sheet['Q' + str(row)].value),
               "auxContraPartida": str(sheet['R' + str(row)].value)
     }

     SumaTotal = sumaAjustes(fila)
     sheet[get_column_letter(column_index_from_string('T')+ row2) + str(row)].value = str(SumaTotal1)

   else:
    fila = {
          "numajuste": str(sheet['A' + str(row)].value),
          "grupo": str(sheet['O' + str(row)].value)[0:3],
          "subgrupo": str((sheet['E' + str(row)].value).split("-")[0]),
          "requerimiento": str(requerimiento[sheet['F' + str(row)].value]),
          "procedencia": str((sheet['G' + str(row)].value).split("-")[0]),
          "categoria": str((sheet['I' + str(row)].value).split("-")[0]),
          "celda": str(ajustes[str(sheet['A' + str(row)].value)]),
          "estado": str((sheet['H' + str(row)].value).split("-")[0]),
          "Partida": str(sheet['N' + str(row)].value),
          "auxPartida": str(sheet['O' + str(row)].value),
          "ContraPartida": str(sheet['Q' + str(row)].value),
          "auxContraPartida": str(sheet['R' + str(row)].value)
      }
    SumaTotal1 = sumaAjustes(fila)
    sheet['T' + str(row)].value = SumaTotal

 wb.save('Plantilla contable consolidada (solo valores) 28022018.xlsx')


# funcion que retorna verdadero si el ajuste es menor
def esMayorSumaCH(ajuste, estadoNICSP, cuantia ):
    if (ajuste == '11')  and (estadoNICSP <> '3' and (cuantia.strip() == "NA" or cuantia.strip() == "MA")):
        return True
    else:
        return False

def esMenorSumaCH(ajuste, estadoNICSP, cuantia ):
    if (ajuste == '10')  and (estadoNICSP <> '3' and (cuantia.strip() == "NA" or cuantia.strip() == "MA")):
        return True
    else:
        return False

def esMenorSumaDe(ajuste, alternativa, estadoNICSP, cuantia):
    if (( ajuste == '14') and (alternativa <> '2' and estadoNICSP <> '3' and cuantia.strip() == "NA" or cuantia.strip() == "MA")):
        return True
    else:
        return False

def esMayorSumaDe(ajuste, alternativa, estadoNICSP, cuantia):
    if (( ajuste == '15') and (alternativa <> '2' and estadoNICSP <> '3' and cuantia.strip() == "NA" or cuantia.strip() == "MA")):
        return True
    else:
        return False


def esBajaSuma(estadoRCP, cuantia, ajuste):
    if ((cuantia.strip() == "MI" or cuantia.strip() == "ME") and (((estadoRCP <> '5') and (estadoRCP <> '8')) and (int(ajuste)) < 6)):
        return True
    else:
        return False


def esEliminacionDepreSuma(ajuste, alternativa, cuantia, estadoRCP):
    if (((alternativa == '2' and ajuste == '6') and (cuantia.strip() == "MA" or cuantia.strip() == "NA")) or ((alternativa == '2' and ajuste == '6') and (((estadoRCP == '5') or (estadoRCP == '8'))))):
        return True
    else:
        return False


def esEliminacionProviValoDesSuma(estadoRCP, estadoNICSP, cuantia, ajuste):

    if ((int(ajuste) >= 7 and (int(ajuste) <= 9) and (((estadoRCP == '5') or (estadoRCP == '8')))) or ( int(ajuste) >= 7 and (int(ajuste) <= 9) and (cuantia.strip() == "MA" or cuantia.strip() == "NA"))):
        return True
    else:
        return False


# funcion que recorre la hoja ByS y hace la sumatoria de los campos segun el ajuste
def sumaAjustes(fila):
    sheet = wb['hoja_de_trabajo_bys']
    vAjuste = 0
    for row in range(2, sheet.max_row + 1):
     if (sheet['B' + str(row)].value == fila["requerimiento"] and sheet['F' + str(row)].value == fila["grupo"] and sheet['H' + str(row)].value == fila["subgrupo"] and sheet['L' + str(row)].value == fila["procedencia"] and sheet['Q' + str(row)].value == fila["estado"] and sheet['N' + str(row)].value == str(fila["categoria"])):
      if (sheet[fila["celda"] + str(row)].value <> '0' and sheet[fila["celda"] + str(row)].value is not None):
       if (esMayorSumaDe(fila["numajuste"], sheet['S' + str(row)].value, sheet['R' + str(row)].value, sheet['P' + str(row)].value) and float(sheet[fila["celda"] + str(row)].value) > 0):  # mayor
        vAjuste += abs(float(sheet[fila["celda"] + str(row)].value))
        #print "mayorDE"
       elif (esMenorSumaDe(fila["numajuste"], sheet['S' + str(row)].value, sheet['R' + str(row)].value, sheet['P' + str(row)].value) and float(sheet[fila["celda"] + str(row)].value) < 0):  # menor
        vAjuste += abs(float(sheet[fila["celda"] + str(row)].value))
        #print "menorDE"
       elif(esMenorSumaCH(fila["numajuste"], sheet['R' + str(row)].value, sheet['P' + str(row)].value) and float(sheet[fila["celda"] + str(row)].value) < 0):  # menor
        vAjuste += abs(float(sheet[fila["celda"] + str(row)].value))
        #print "menorCH"
       elif(esMayorSumaCH(fila["numajuste"], sheet['R' + str(row)].value, sheet['P' + str(row)].value) and float(sheet[fila["celda"] + str(row)].value) > 0):  # mayor
        vAjuste += abs(float(sheet[fila["celda"] + str(row)].value) )
        #print "mayorCH"
       elif (esBajaSuma(sheet['Q' + str(row)].value, sheet['P' + str(row)].value, fila["numajuste"])):  # bajas
        vAjuste += abs(float(sheet[fila["celda"] + str(row)].value))
        #print "baja"
       elif (esEliminacionDepreSuma(fila["numajuste"], sheet['S' + str(row)].value, sheet['P' + str(row)].value, sheet['Q' + str(row)].value)):  # depreciacion
        vAjuste += abs(float(sheet[fila["celda"] + str(row)].value))
        #print 'elide'
       elif (esEliminacionProviValoDesSuma(sheet['Q' + str(row)].value, sheet['R' + str(row)].value, sheet['P' + str(row)].value, fila["numajuste"]) and float(sheet[fila["celda"] + str(row)].value) <> 0 ):  # Eliminacion Valorizacion Desvalorizacion y provision
        vAjuste += abs(float(sheet[fila["celda"] + str(row)].value))


    return vAjuste

recorrer()

#print sumaAjustes(fila1)
#print sumaAjustes(fila2)
