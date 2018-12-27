#! python 2.7
# -*- coding: cp1252 -*-
# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils import column_index_from_string

# variables de entrada
wb = openpyxl.load_workbook('PlantillaContable28022018.xlsx')
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
                 '2-FONDOS ESPECIALES': '4010',
                 '6-UNISALUD': '4002'}


def muestra_valores(fila):
    print('\n*************************************************')
    valores = fila.items()
    for valor, dato in valores:
        print(valor, '-->', dato)
    print('\n*************************************************')

 fila = {
            "numajuste": str(sheet['A' + str(row)].value),
            "grupo": str(sheet['O' + str(row)].value)[0:3],
            "subgrupo": str((sheet['E' + str(row)].value).split("-")[0]),
            "requerimiento": str(requerimiento[sheet['F' + str(row)].value]),
            "procedencia": str((sheet['G' + str(row)].value).split("-")[0]),
            "categoria": str(sheet['I' + str(row)].value),
            "celda": str(ajustes[str(sheet['A' + str(row)].value)]),
            "estado": str((sheet['H' + str(row)].value).split("-")[0]),
            "Partida": str(sheet['N' + str(row)].value),
            "auxPartida": str(sheet['O' + str(row)].value),
            "ContraPartida": str(sheet['Q' + str(row)].value),
            "auxContraPartida": str(sheet['R' + str(row)].value)
        }

muestra_valores(fila)

def recorrer():
    for row in range(2, sheet.max_row + 1):
        fila = {

            "numajuste": str(sheet['A' + str(row)].value),
            "grupo": str(sheet['O' + str(row)].value)[0:3],
            "subgrupo": str((sheet['E' + str(row)].value).split("-")[0]),
            "requerimiento": str(requerimiento[sheet['F' + str(row)].value]),
            "procedencia": str((sheet['G' + str(row)].value).split("-")[0]),
            "categoria": str(sheet['I' + str(row)].value),
            "celda": str(ajustes[str(sheet['A' + str(row)].value)]),
            "estado": str((sheet['H' + str(row)].value).split("-")[0]),
            "Partida": str(sheet['N' + str(row)].value),
            "auxPartida": str(sheet['O' + str(row)].value),
            "ContraPartida": str(sheet['Q' + str(row)].value),
            "auxContraPartida": str(sheet['R' + str(row)].value)
        }

        validaAux(fila)


# ----------------------------------------------------------------------

# funcion que retorna verdadero si el ajuste es menor
def esMayorSumaCH(ajuste, estadoNICSP, cuantia):
    if (ajuste == '11') and (estadoNICSP <> '3' and (cuantia.strip() == "NA" or cuantia.strip() == "MA")):
        return True
    else:
        return False


def esMenorSumaCH(ajuste, estadoNICSP, cuantia):
    if (ajuste == '10') and (estadoNICSP <> '3' and (cuantia.strip() == "NA" or cuantia.strip() == "MA")):
        return True
    else:
        return False


def esMenorSumaDe(ajuste, alternativa, estadoNICSP, cuantia):
    if ((ajuste == '14') and (
            alternativa <> '2' and estadoNICSP <> '3' and cuantia.strip() == "NA" or cuantia.strip() == "MA")):
        return True
    else:
        return False


def esMayorSumaDe(ajuste, alternativa, estadoNICSP, cuantia):
    if ((ajuste == '15') and (
            alternativa <> '2' and estadoNICSP <> '3' and cuantia.strip() == "NA" or cuantia.strip() == "MA")):
        return True
    else:
        return False


def esBajaSuma(estadoRCP, cuantia, ajuste):
    if ((cuantia.strip() == "MI" or cuantia.strip() == "ME") and (
            ((estadoRCP <> '5') and (estadoRCP <> '8')) and (int(ajuste)) < 6)):
        return True
    else:
        return False


def esEliminacionDepreSuma(ajuste, alternativa, cuantia, estadoRCP):
    if (((alternativa == '2' and ajuste == '6') and (cuantia.strip() == "MA" or cuantia.strip() == "NA")) or (
            (alternativa == '2' and ajuste == '6') and (((estadoRCP == '5') or (estadoRCP == '8'))))):
        return True
    else:
        return False


def esEliminacionProviValoDesSuma(estadoRCP, estadoNICSP, cuantia, ajuste):
    if ((int(ajuste) >= 7 and (int(ajuste) <= 9) and (((estadoRCP == '5') or (estadoRCP == '8')))) or (
            int(ajuste) >= 7 and (int(ajuste) <= 9) and (cuantia.strip() == "MA" or cuantia.strip() == "NA"))):
        return True
    else:
        return False

#----------------------------------------------------------------------------------------------------------------------
# funcion que valida auxiliares y cuentas
def validaAux(fila):
    sheet = wb['hoja_de_trabajo_bys_19']
    tipo = ""

    for row in range(2, sheet.max_row + 1):
     if (sheet['B' + str(row)].value == fila["requerimiento"] and sheet['F' + str(row)].value == fila["grupo"] and sheet['H' + str(row)].value == fila["subgrupo"] and sheet['L' + str(row)].value == fila["procedencia"] and sheet['Q' + str(row)].value == fila["estado"] and sheet['N' + str(row)].value == str(fila["categoria"])):  # Compara filtros
      if (sheet[fila["celda"] + str(row)].value <> '0' and sheet[fila["celda"] + str(row)].value is not None):  # Valida que exista un ajuste en bienes y servicios
       if (esMayorSumaCH(fila["numajuste"], sheet['R' + str(row)].value, sheet['P' + str(row)].value) and float(sheet[fila["celda"] + str(row)].value) > 0):
        if (sheet[get_column_letter(column_index_from_string(fila["celda"]) + 2) + str(row)].value <> str(fila["auxPartida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 5) + str(row)].value <> str(fila["auxContraPartida"]) and sheet[get_column_letter(column_index_from_string(fila["celda"]) + 1) + str(row)].value <> str(fila["Partida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 4) + str(row)].value <> str(fila["ContraPartida"])):  # Valida cuentas de costo hisotrico y depreciacion MAYOR
         global contador
         contador += 1
         muestra_valores(fila)
         tipo = "esMayorSumaCH"
         print tipo
         return False

       elif (esMenorSumaCH(fila["numajuste"], sheet['R' + str(row)].value, sheet['P' + str(row)].value) and float(sheet[fila["celda"] + str(row)].value) < 0):
        if (sheet[get_column_letter(column_index_from_string(fila["celda"]) + 2) + str(row)].value <> str(fila["auxPartida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 5) + str(row)].value <> str(fila["auxContraPartida"]) and sheet[get_column_letter(column_index_from_string(fila["celda"]) + 1) + str(row)].value <> str(fila["Partida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 4) + str(row)].value <> str(fila["ContraPartida"])):  # Valida cuentas de costo hisotrico y depreciacion menor
         global contador
         contador += 1
         muestra_valores(fila)
         tipo = "esMenorSumaCH"
         print tipo
         return False

       elif (esMenorSumaDe(fila["numajuste"], sheet['S' + str(row)].value, sheet['R' + str(row)].value, sheet['P' + str(row)].value) and float(sheet[fila["celda"] + str(row)].value) < 0):
        if (sheet[get_column_letter(column_index_from_string(fila["celda"]) + 2) + str(row)].value <> str(fila["auxPartida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 5) + str(row)].value <> str(fila["auxContraPartida"]) and sheet[get_column_letter(column_index_from_string(fila["celda"]) + 1) + str(row)].value <> str(fila["Partida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 4) + str(row)].value <> str(fila["ContraPartida"])):  # Valida cuentas de costo hisotrico y depreciacion menor
         global contador
         contador += 1
         muestra_valores(fila)
         tipo = "esMenorSumaDe"
         print tipo
         return False

       elif (esMayorSumaDe(fila["numajuste"], sheet['S' + str(row)].value, sheet['R' + str(row)].value,sheet['P' + str(row)].value) and float(sheet[fila["celda"] + str(row)].value) > 0):
        if (sheet[get_column_letter(column_index_from_string(fila["celda"]) + 2) + str(row)].value <> str(fila["auxPartida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 5) + str(row)].value <> str(fila["auxContraPartida"]) and sheet[get_column_letter(column_index_from_string(fila["celda"]) + 1) + str(row)].value <> str(fila["Partida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 4) + str(row)].value <> str(fila["ContraPartida"])):  # Valida cuentas de costo hisotrico y depreciacion menor
         global contador
         contador += 1
         muestra_valores(fila)
         tipo = "esMayorSumaDe"
         print tipo
         return False

       elif (esBajaSuma(sheet['Q' + str(row)].value, sheet['P' + str(row)].value, fila["numajuste"])):
        if (sheet[get_column_letter(column_index_from_string(fila["celda"]) + 2) + str(row)].value <> str(fila["auxPartida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 5) + str(row)].value <> str(fila["auxContraPartida"]) and sheet[get_column_letter(column_index_from_string(fila["celda"]) + 1) + str(row)].value <> str(fila["Partida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 4) + str(row)].value <> str(fila["ContraPartida"])):  # Valida cuentas de costo hisotrico y depreciacion menor
         global contador
         contador += 1
         muestra_valores(fila)
         tipo = "esBajaSuma"
         print tipo
         return False

       elif ( esEliminacionDepreSuma(fila["numajuste"], sheet['S' + str(row)].value, sheet['P' + str(row)].value, sheet['Q' + str(row)].value)):
        if (sheet[get_column_letter(column_index_from_string(fila["celda"]) + 2) + str(row)].value <> str(fila["auxPartida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 5) + str(row)].value <> str(fila["auxContraPartida"]) and sheet[get_column_letter(column_index_from_string(fila["celda"]) + 1) + str(row)].value <> str(fila["Partida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 4) + str(row)].value <> str(fila["ContraPartida"])):  # Valida cuentas de costo hisotrico y depreciacion menor
         global contador
         contador += 1
         muestra_valores(fila)
         tipo = "esEliminacionDepreSuma"
         print tipo
         return False

       elif (esEliminacionProviValoDesSuma(sheet['Q' + str(row)].value, sheet['R' + str(row)].value, sheet['P' + str(row)].value, fila["numajuste"]) and float(sheet[fila["celda"] + str(row)].value) <> 0):
        if (sheet[get_column_letter(column_index_from_string(fila["celda"]) + 2) + str(row)].value <> str(fila["auxPartida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 5) + str(row)].value <> str(fila["auxContraPartida"]) and sheet[get_column_letter(column_index_from_string(fila["celda"]) + 1) + str(row)].value <> str(fila["Partida"]) or sheet[get_column_letter(column_index_from_string(fila["celda"]) + 4) + str(row)].value <> str(fila["ContraPartida"])):  # Valida cuentas de costo hisotrico y depreciacion menor
         global contador
         contador += 1
         muestra_valores(fila)
         tipo = "esEliminacionProviValoDesSuma"
         print tipo
         return False

    return True


#recorrer()
print contador


