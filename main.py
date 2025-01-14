import xlsxwriter
import json

def cargar_json():
    global data
    with open('notes.json', 'r') as file:
        data = json.load(file)

filename = "exercici0.xlsx"
data = None
cargar_json()

workbook = xlsxwriter.Workbook(filename)
worksheet0 = workbook.add_worksheet('Full 0')
worksheet1 = workbook.add_worksheet('Full 1')

# ESTILOSS
red_background = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF', 'align': 'center'})
green_background = workbook.add_format({'bg_color': '#20fc03', 'font_color': '#000000', 'align': 'center'})

def calcularMitja(info):
    percentatges = [0.1, 0.1, 0.1, 0.2, 0.5]
    notes = [info["PR01"], info["PR02"], info["PR03"], info["PR04"], info["EX01"]]
    notaFinal = sum(notes[i] * percentatges[i] for i in range(len(notes)))
    return notaFinal

def crear_hoja(worksheet, noms):
    #Possar les capçaleres amb el num de la cassella
    if noms == "Nom":
        worksheet.write(0, 0, "Nom")
    else:
        worksheet.write(0, 0, "Identificació")
        worksheet.write(0, 1, 'PR01')
        worksheet.write(0, 2, 'PR02')
        worksheet.write(0, 3, 'PR03')
        worksheet.write(0, 4, 'PR04')
        worksheet.write(0, 5, 'EX01')
        worksheet.write(0, 6, '%Faltes')
        worksheet.write(0, 7, 'Vàlid')
        worksheet.write(0, 8, 'Total')

    for x, alumne in enumerate(data):
        fila = x + 1
        if noms == "Nom":
            worksheet.write(fila, 0, alumne["Name"])
        else:
            worksheet.write(fila, 0, alumne["id"][1:5])

        notes = [alumne["PR01"], alumne["PR02"], alumne["PR03"], alumne["PR04"], alumne["EX01"]]
        for col, nota in enumerate(notes, 1):
            worksheet.write(fila, col, nota)
        
        worksheet.write(fila, 6, alumne["%Faltes"])

        if alumne["%Faltes"] <= 20 and alumne["EX01"] >= 4:
            worksheet.write(fila, 7, "true")
            mitja = calcularMitja(alumne)
        else:
            worksheet.write(fila, 7, "false")
            mitja = 1

        worksheet.write(fila, 8, mitja)
        
    worksheet.conditional_format(f'B2:F{len(data) + 1}', {
        'type': 'cell',
        'criteria': '<',
        'value': 5,
        'format': red_background
    })
    worksheet.conditional_format(f'I2:I{len(data) + 1}', {
        'type': 'cell',
        'criteria': '<',
        'value': 5,
        'format': red_background
    })
    worksheet.conditional_format(f'I2:I{len(data) + 1}', {
        'type': 'cell',
        'criteria': '>=',
        'value': 7,
        'format': green_background
    })

# Creación de las tablas
crear_hoja(worksheet0, "Nom")
crear_hoja(worksheet1, "Id")

workbook.close()
print(f"Generated: '{filename}'")
