import openpyxl
import argparse

def main(book_name):
    # Cargar el libro de Excel
    book = openpyxl.load_workbook('./Data/GlobalDataUpdated (version 2) 05-11-2024.xlsx')

    # Seleccionar la hoja activa
    sheet = book.active
    string = """
Wacom
AWS
Master Card
Lineea DataScan
BTG Pactual 
Coppel
Avante group SAS
ISSQUARED Inc
Lennox
Samsung 
Circulo Corp
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
