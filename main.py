import openpyxl
import argparse
from config import INVALID_VALUES
from utils import clean_value, ensure_columns_exist, format_contact, insert_name_columns, normalize_text, normalize_phone, normalize_interlocutor, format_opportunity, set_amount, set_campaign, set_closing_date, set_currency, set_flow, set_owner, set_stage, set_state, set_tags

def main(name, flow, owner, date):
    input_file = f"./Clean/{name}ForClean.xlsx"
    output_file = f"./Clientify/{name}ForClientify.xlsx"
    
    book = openpyxl.load_workbook(input_file)
    sheet = book.active

    ensure_columns_exist(sheet)
    insert_name_columns(sheet)

    column_names = {col[0].value: col for col in sheet.columns}

    for column in sheet.columns:
        if column[0].value in INVALID_VALUES:
            continue

        if column[0].value in ["NOMBRE COMERCIAL EMPRESA", "PAÍS"]:
            for cell in column:
                normalize_text(cell)

        elif column[0].value in ["TELÉFONO DE EMPRESA", "TELÉFONO EMPRESA", "TELEFONO 1", "TELEFONO 2", "CELULAR", ]:
            for cell in column:
                normalize_phone(cell, sheet)

        elif column[0].value == "INTERLOCUTOR 1":
            for cell in column:
                normalize_interlocutor(cell, sheet)

        elif column[0].value == "OPORTUNIDAD":
            for cell in column[1:]:  # Skip the first cell
                format_opportunity(cell, sheet, flow)

        elif column[0].value == "CONTACTO":
            for cell in column[1:]:  # Skip the first cell
                format_contact(cell, sheet)
                
                
        for cell in column:
            clean_value(cell, INVALID_VALUES)

    set_campaign(sheet, name)
    set_tags(sheet, name)
    set_amount(sheet )
    set_currency(sheet)
    set_state(sheet)
    set_stage(sheet, name)
    set_flow(sheet, name, flow)
    set_owner(sheet, owner)
    set_closing_date(sheet, date)


    book.save(output_file)
    print(f"File saved as {output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process an Excel file.')
    parser.add_argument('name', type=str, help='The name to be used in the filenames.')
    parser.add_argument('flow', type=str, help='The flow to be used in the opportunity formatting.')
    parser.add_argument('owner', type=str, help='The owner to be used in the column PROPIETARIO.')
    parser.add_argument('date', type=str, help='The date to be used in the column FECHA CIERRE.')
    args = parser.parse_args()
    main(args.name, args.flow, args.owner, args.date)
