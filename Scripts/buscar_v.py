import openpyxl
import argparse

def main(book_name):
    # Cargar el libro de Excel
    book = openpyxl.load_workbook('./Data/GlobalDataUpdated (version 1).xlsx')

    # Seleccionar la hoja activa
    sheet = book.active
    string = """
3M
AB InBev
Abastible
ABT
ACHS
Adidas
AFEX Transferencias y Cambios
Agrichile
AGROSUPER
Aguas Andinas
Aguas Antofagasta
Aguas Araucanía
Ahorrame
Ahorrocoop
Algramo
Ambrosoli
Andes Industrial Ltda.
Anida Latam
Antofagasta Minerals
Aramark
Arcoprime
Ariztia
Atlas
Avianca
AVLA
Axteroid
Babytuto
Banchile inversiones
Banco BICE
Banco BTG Pactual Chile
Banco Citi
Banco de Chile
Banco de crédito e inversiones - BCI
Banco de la Nación Argentina
Banco do Brasil
Banco Edwards
Banco Falabella
Banco Internacional
Banco Nova
Banco Santander-Chile
Banco Security
Banefe
Bank of China
BICE Vida
BK
BRINK'S
Btrust
Burea
Caja Los Andes
Caja los Héroes
Canal 13
Cannon Home
Capual
Carozzi
Casa Royal
Castaño
CCLA
CCU
Celmedia SPA
Celulosa Arauco
CGE
ChileMat
Chilevisión
Chilquinta
Cimenta
ClaroVTR
CLOROX
CMPC
CNN Chile
Cocha
Cochilco
Codelco
Colgate-palmolive
COLUN
Compass
Concha y Toro
Consalud
Coocretal
Copec
COPESA
Cornershop
Corporacion RM
CrediChile
Cristal
Cruz verde
Decathlon
Derco
Detacoop
E-Digital
Easy
EL MERCURIO
Emasa
Entel Connect Center
ESTEC
Falabella
FALP
Familyshop
Farmacias Ahumada
Femsa
FiberHome Chile
FID Seguros
FINTUAL
FLOW SA
Forus
Fpay
GDExpress
GERDAU AZA
Gesintel
GLENCORE
GREEN GLASS
Grupo CAP
Grupo Security
GRUPO SECURITY
GTD
HDI
Help seguros
Hites
HKN
Holding Altas Cumbres
HSBC Bank (Chile)
IANSA
IDC Chile
IDEAL SA
IKEA
InnovaWeb
Interexport
Intergas
Ionix - Zeleri
IPX DIGITAL
Itaú CorpBanca
Johnson's
Jumbo
Karibu
Karun
Kimberly Clark
Klap
Kunstmann
L'Oreal
La positiva Seguros
La red
LABLAB
LaPolar
LarrainVial
Latam Airlines
Ledrium
Licita Pyme
Líder
Lipigas
Mallplaza
Mar del Sur SpA
Massiva
Melón
Mercado Libre
Metbus
MetLife
MUFG Bank
Mundialis SpA
Mundo
Mutual de Seguridad
Natura
Negocio Financiero Solventa Tarjetas
Nestlé
Netline
NIKE
NSAgro
Openagora
Oriencoop
Overall
OXXO Chile
Paris (Cencosud)
Parque Arauco
Patagonia
PepsiCo
PF
Pichara
Polla chilena
PRIMUS CAPITAL
Procter and Gamble
Proexsi
Prosud SA
Pullman Bus
Puma Chile SpA
Quantum
Quiñenco
Reale seguros
Redbus
Reservo
Ripley
Rosen
Salcobrand
Samsonite
Samsung
Santander Chile
Scotia Cencosud
Scotiabank Azul (BBVA)
Shop Chile
SIGDO KOPPERS
Sixbell
SKY
SMU
Sodexo
Sodimac
Soprole
STARBUCKS
STNG
Subcargo
Súper Pollo
Sura
TBanc
Telefonica Chile Ltda
Televisión Nacional de Chile
Telsur
The Blue Commerce
The Wild Brand
TIKA
TOC-TOC
Tommy & Calvin Klain
Tottus
Transbank
Transvip
Transytec
Tresmontes Lucchetti
Trionip Spa
Turbus
TVN Chile
UCV Televisión
Unilever
Unimarc
Vida Camara Seguros
VIÑA EMILIANA
VIÑA MONTES
VIÑA SAN PEDRO
VIÑA SANTA RITA
Virgin Mobile
Walmart
WATTS
WOM
Zurich Santander
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
