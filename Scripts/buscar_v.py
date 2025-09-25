import openpyxl
import argparse

def main(book_name):
    # Cargar el libro de Excel
    book = openpyxl.load_workbook('./Data/GlobalDataUpdated13-07-2025.xlsx')

    # Seleccionar la hoja activa
    sheet = book.active
    string = """Vivo Corp
Workmate
Automotriz Salfa Sur Ltda
Reutter S.A.
Supermercado del Neumatico Ltda.
Empresas Melón
Melón
Valvulas Industriales S.A.
Clinical Market s.a.
Epysa Implementos Ltda.
VETO Y CIA LTDA.
Tecno Fast
Vina Santa Carolina Sa Santiago Cl
Mora Pavic Odontología
Club Providencia
BENEO
Tucapel
Viña De Martino
Vitel Energia
UCMChile - Unidad Coronaria Móvil
TECNOGLOBAL
ALO GROUP
ARRIMAQ
GRUPO PESCO
GRUPO SIMMA
JANSSEN SA
KOMATSU REMAN CENTER CHILE
KSB CHILE SA
AKVA group Chile
PRECISION
Samsung Electronics Chile
Laboratorio Chile | Teva
Tecnofarma Sa
Compañia Chilena de Fosforos S.A.
Mainstream Renewable Power
Automotores Gildemeister SpA
Caren SpA
Isa Intervial
Moneda Asset
Moneda Asset Management
Multicaja Sa
Uno Afp
Colbun S.A
Synthon
Caffarena
Artel Sa
Bash Administracion Limitada
Cia. Industrial El Volcan S.A.
Cosmoplas S.A.
CAP S.A.
"""

    # Lista de empresas
    EMPRESAS = string.split("\n")
    # Limpiar y normalizar los nombres de las empresas
    EMPRESAS_CLEAN = [x.upper().strip().replace("  ", " ").replace(",", "").replace(".", "").replace("(", "").replace(")", "") for x in EMPRESAS]

    # Crear una nueva hoja para los datos filtrados
    new_sheet = book.create_sheet("Filtrados")

    # Buscar la columna "NOMBRE COMERCIAL EMPRESA"
    col_index = None
    for col in sheet.iter_cols(1, sheet.max_column):
        # if col[0].value == "URL LINKEDIN EMPRESA": 
        if col[0].value == "NOMBRE COMERCIAL EMPRESA":
            col_index = col[0].column
            break

    if col_index:
        # Copiar el encabezado a la nueva hoja
        for cell in sheet[1]:
            new_sheet[cell.coordinate] = cell.value

        new_row_index = 2  # Comenzar desde la segunda fila para los datos
        # Iterar sobre las filas y copiar las que cumplen con el criterio
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
            cell_value = row[col_index-1].value
            if cell_value and cell_value.upper().strip().replace("  ", " ").replace(",", "").replace(".", "").replace("(", "").replace(")", "") in EMPRESAS_CLEAN:
                
                for cell in row:
                    new_sheet.cell(row=new_row_index, column=cell.column, value=cell.value)
                new_row_index += 1

    # Eliminar la hoja original si lo deseas
    book.remove(sheet)

    # Guardar el archivo modificado
    book.save(f'./Extraídos/{book_name}.xlsx')
    for empresa in EMPRESAS_CLEAN:
        print(empresa)
    print(f"Archivo guardado como '{book_name}.xlsx'")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Filtrar datos de un archivo de Excel y guardarlos en un nuevo archivo.")
    parser.add_argument("book_name", help="Nombre del archivo de salida (sin extensión)")
    args = parser.parse_args()
    main(args.book_name)