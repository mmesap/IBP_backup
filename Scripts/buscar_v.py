import openpyxl
import argparse

def main(book_name):
    # Cargar el libro de Excel
    book = openpyxl.load_workbook('./Data/GlobalDataUpdated-6-9-2025.xlsx')

    # Seleccionar la hoja activa
    sheet = book.active
    string = """
AGT TECNOLOGIAS VE, S.A.
Laboratorios Elmor, S.A.
Oxiteno S.A.
3PL Panamericana, C.A.
MT2005
Bancamiga Banco de Desarrollo
FV & ASOCIADOS CA
Setecsa de Venezuela
APT TECNOLOGIA Y SISTEMAS C.A
INDUCHEM C A
GRUPO PHX C A
PROAGRO PROTINAL
FUNDACION VENEZOLANO - ALEMANA COLEGIO HUMBOLDT
CONSORCIO CREDICARD
FOSPUCA Internacional
INVERSIONES RESANSIL CONSTRUCTORA
IMPORTADORA USY C A
Banco de Comercio Exterior-Bancoex
TICKETMUNDO, CA.
BANESCO SEGUROS, C.A.
VC MEDIOS C.A.
PRINTING SUPPLY INTERNATIONAL C A
CENTRO DIAGNOSTICO DOCENTE LAS MERCEDES C A
GENICA GENERAL DE ALIMENTOS NICOLAS E ISABEL DISTRIBUIDORA G
CONSTRUCTORES DE COMERCIO CAMARGO CORREA S.A.
DVL SERVICIO Y REPRESENTACIONES C A
SEGUROS PIRAMIDE C A
COSMETICOS ROLDA
MG Group, CA.
BOLSA DE VALORES DE CARACAS
FUNDACIÓN OSCARYANNY
Nozomi Salud, Casa de Representacion CA
1000 MOTORSPORT C.A.
IMPREGILO
COMERCIAL GIL S.A.
INDUSTRIAS RUANSA DE VENEZUELA
AVIOR AIRLINES C A
OMICRON C A
CASA DE REPRESENTACIONES JMW CA
FARMACIA COLSALUD CA
BANCO PLAZA C A
LUMALAC DAIRY PRODUCT LUMALAC C A
INVERSIONES EMPLEATE,C.A.
Inelectra Venezuela
DEAR C A
Soluciones Netready
World Trading Casa de Bolsa CA
CONSTRUCTORA SAMBIL
EUROBUILDING INTERNACIONAL C A
MINI BRUNO SUCESORES C A
Alimentos La Caridad, C.A.
GRIFOCENTRO C A
ALIMENTOS MUNCHYS C A
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