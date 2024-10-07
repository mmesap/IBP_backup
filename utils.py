from config import AMOUNT, CAMPAIGN, CLOSING_DATE, CURRENCY, FLOW, OWNER, REQUIRED_COLUMNS, STAGE, STATE, TAGS, FIRST_NAME_COLUMN, LAST_NAME_COLUMN, NAME_COLUMN

def clean_value(cell, invalid_values):
    if cell.value in invalid_values:
        cell.value = None

def normalize_text(cell):
    if cell.value is not None:
        cell.value = str(cell.value).upper().strip()
        cell.value = cell.value.replace("  ", " ").replace(",", "").replace(".", "").replace("(", "").replace(")", "")

def normalize_phone(cell, sheet):
    if cell.value is not None:
        cell_value = str(cell.value).replace("+", "").strip().replace(" ", "").replace(".", "").replace("(", "").replace(")", "").replace("-", "").replace(",", "/").replace("PBX:", "").replace("Y", "/").replace("y", "/")
        phones = cell_value.split("/")
        cell.value = "+" + phones[0]
        if len(phones) > 1:
            insert_additional_phone_column(cell, sheet, phones[1])

def insert_additional_phone_column(cell, sheet, additional_phone):
    col_idx = cell.column + 1
    if sheet.cell(row=1, column=col_idx).value != sheet.cell(row=1, column=cell.column).value + " OTRO":
        sheet.insert_cols(col_idx, 1)
        sheet.cell(row=1, column=col_idx).value = sheet.cell(row=1, column=cell.column).value + " OTRO"
    sheet.cell(row=cell.row, column=col_idx).value = f"+{additional_phone}"

def normalize_interlocutor(cell, sheet):
    if cell.value:
        names = cell.value.strip().split(" ")
        if len(names) > 2:
            cell.value = f"{names[0]} {names[1]}"
            sheet.cell(row=cell.row, column=cell.column + 1).value = " ".join(names[2:])
        elif len(names) == 2:
            cell.value = names[0]
            sheet.cell(row=cell.row, column=cell.column + 1).value = names[1]
    sheet.cell(row=1, column=cell.column).value = "NOMBRES"
    sheet.cell(row=1, column=cell.column + 1).value = "APELLIDOS"
    sheet.cell(row=1, column=cell.column + 2).value = "OPORTUNIDAD"

def format_opportunity(cell, sheet, flow):
    company_name = get_column_value(sheet, cell.row, NAME_COLUMN)
    first_name = get_column_value(sheet, cell.row, FIRST_NAME_COLUMN)
    last_name = get_column_value(sheet, cell.row, LAST_NAME_COLUMN)
    cell.value = f"{company_name} / {flow} / {first_name} {last_name}"

def format_contact(cell, sheet):
    first_name = get_column_value(sheet, cell.row, FIRST_NAME_COLUMN)
    last_name = get_column_value(sheet, cell.row, LAST_NAME_COLUMN)
    cell.value = f"{first_name} {last_name}"

def get_column_value(sheet, row, column_header):
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == column_header:
            return col[row-1].value
    return ""

def set_campaign(sheet, name):
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == CAMPAIGN:
            for cell in col[1:]:  # Skip the first cell (header)
                cell.value = name
            break
def set_tags(sheet, name):
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == TAGS:
            for cell in col[1:]:  # Skip the first cell (header)
                cell.value = name
            break

def set_amount(sheet):
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == AMOUNT:
            for cell in col[1:]:  # Skip the first cell (header)
                cell.value = "1"
            break

def set_currency(sheet):
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == CURRENCY:
            for cell in col[1:]:  # Skip the first cell (header)
                cell.value = "USD"
            break

def set_state(sheet):
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == STATE:
            for cell in col[1:]:  # Skip the first cell (header)
                cell.value = "Abierta"
            break

def set_stage(sheet, name):
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == STAGE:
            for cell in col[1:]:  # Skip the first cell (header)
                cell.value = f"LEAD SIN GESTIÓN - {name.upper()}"
            break
def set_flow(sheet, name, flow):
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == FLOW:
            for cell in col[1:]:  # Skip the first cell (header)
                cell.value = f"{name.upper()} {flow.upper()}"
            break

def set_owner(sheet, owner):
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == OWNER:
            for cell in col[1:]:  # Skip the first cell (header)
                cell.value = owner
            break
def set_closing_date(sheet, date):
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == CLOSING_DATE:
            for cell in col[1:]:  # Skip the first cell (header)
                cell.value = date
            break
        
def ensure_columns_exist(sheet):
    existing_columns = [cell.value for cell in sheet[1]]
    for column in REQUIRED_COLUMNS:
        if column not in existing_columns:
            sheet.insert_cols(len(existing_columns) + 1)
            sheet.cell(row=1, column=len(existing_columns) + 1).value = column
            existing_columns.append(column)

def insert_name_columns(sheet):
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == "INTERLOCUTOR 1":
            insert_idx = col[0].column + 1
            sheet.insert_cols(insert_idx, 3)  # Insert three columns
            sheet.cell(row=1, column=insert_idx).value = "NOMBRES"
            sheet.cell(row=1, column=insert_idx + 1).value = "APELLIDOS"
            sheet.cell(row=1, column=insert_idx + 2).value = "CONTACTO"
            break
