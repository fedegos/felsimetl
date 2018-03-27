import openpyxl
import time
import json
from datetime import datetime
from constants import *

from tableextraction.tableunpacker import *
from categoriesreader.categoriesreader import *
from exceladapter.excelreader import *
from exceladapter.excelwriter import *

from translators.googletranslator import *

# abrir

categories_reader = CategoriesReader()
actual_excelreader = ExcelReader('inputs/PLANILLA CONTABILIDAD ACTIVA 2017.xlsx')
current_accounts_excelreader = ExcelReader('inputs/CUENTAS CORRIENTES.xlsx')

cheques_sheet = actual_excelreader.get_sheet(CHEQUES_SHEET_NAME)
caja_sheet = actual_excelreader.get_sheet(CAJA_SHEET_NAME)
credicoop_sheet = actual_excelreader.get_sheet(CREDICOOP_SHEET_NAME)

cuentas_corrientes_sheet = current_accounts_excelreader.get_sheet(CUENTAS_CORRIENTES_SHEET_NAME)

# initialize lists

actual_flows = []
projected_flows = []
new_categories = set()

for rowNum in range(3, cheques_sheet.max_row):  # skip the first row
    providers_amount = cheques_sheet.cell(row=rowNum, column=6).value

    if not providers_amount:
        continue

    check = {}
    date_value = cheques_sheet.cell(row=rowNum, column=1).value
    details = cheques_sheet.cell(row=rowNum, column=2).value

    check['date'] = date_value.strftime("%d/%m/%Y") if date_value else "N/D"
    check['year'] = date_value.strftime("%Y") if date_value else "N/D"
    check['week'] = int(date_value.strftime("%U")) + 1 if date_value else "N/D"

    check['flow'] = ACTUAL_FLOW
    check['flexibility'] = ""
    check['type'] = CHEQUES_TYPE  # cheques
    check['details'] = details

    category, account = categories_reader.get_category_from_details(details)
    check['category'] = category
    check['account'] = account
    check['income'] = ""
    check['expense'] = to_google_num(providers_amount)

    if date_value and date_value > datetime.now():
        check['flow'] = ESTIMATED_FLOW
        check['flexibility'] = INFLEXIBLE
        print(check)
        projected_flows.append(check)
    else:
        actual_flows.append(check)

table_unpacker = TableUnpacker(caja_sheet)

for rowNum in range(3, caja_sheet.max_row):
    cash_flow = {}

    row_unpacker = table_unpacker.get_row_unpacker(rowNum)

    account = row_unpacker.get_value_at(3)
    details = row_unpacker.get_value_at(2)

    if details == CARGAS_SOCIALES:
        account = SUELDOS

    if details in BANKS:
        account = PRESTAMOS_BANCARIOS

    if (account == T_E_CUENTAS_PROPIAS and details == FELSIM_CREDICOOP):
        continue

    expense = row_unpacker.get_value_at(5)
    income = row_unpacker.get_value_at(6)
    category = categories_reader.get_category_from_account_and_details(account, details)

    if category == NEW_CATEGORY:
        new_categories.add('{"account": "%s", "details": "%s"}' % (account, details))

    date_value = row_unpacker.get_value_at(1)

    cash_flow['date'], cash_flow['week'], cash_flow['year'] = unpack_dates(date_value)

    cash_flow['flow'] = ACTUAL_FLOW
    cash_flow['flexibility'] = ""

    cash_flow['type'] = CAJAS_TYPE  # caja

    cash_flow['details'] = details

    cash_flow['category'] = category
    cash_flow['account'] = account
    cash_flow['income'] = ""
    cash_flow['expense'] = to_google_num(expense) if expense else ""
    cash_flow['income'] = to_google_num(income) if income else ""

    # print(cash_flow)

    actual_flows.append(cash_flow)

credicoop_unpacker = TableUnpacker(credicoop_sheet)

for rowNum in range(3, credicoop_sheet.max_row):
    cash_flow = {}

    row_unpacker = credicoop_unpacker.get_row_unpacker(rowNum)

    details = row_unpacker.get_value_at(6)
    account = row_unpacker.get_value_at(7)

    if details == CARGAS_SOCIALES:
        account = SUELDOS

    if details in BANKS:
        account = PRESTAMOS_BANCARIOS

    if details == ANNULED:
        continue

    if account == T_E_CUENTAS_PROPIAS and details == "FELSIM CAJA":
        continue

    if details.startswith(MORATORIA_AFIP):
        account = MORATORIAS

    expense = row_unpacker.get_value_at(9)
    income = row_unpacker.get_value_at(10)
    category = categories_reader.get_category_from_account_and_details(account, details)

    if category == NEW_CATEGORY:
        new_categories.add('{"account": "%s", "details": "%s"}' % (account, details))

    date_value = row_unpacker.get_value_at(1)
    check_clearing_date_value = row_unpacker.get_value_at(4)

    if check_clearing_date_value:
        date_value = check_clearing_date_value

    cash_flow['date'], cash_flow['week'], cash_flow['year'] = unpack_dates(date_value)

    cash_flow['flow'] = ACTUAL_FLOW
    cash_flow['flexibility'] = ""
    cash_flow['type'] = CREDICOOP_TYPE
    cash_flow['details'] = details

    cash_flow['category'] = category
    cash_flow['account'] = account
    cash_flow['income'] = ""
    cash_flow['expense'] = to_google_num(expense) if expense else ""
    cash_flow['income'] = to_google_num(income) if income else ""

    # print(cash_flow)

    actual_flows.append(cash_flow)

# crear excel nuevo
filename = 'outputs/consolidado_real_' + time.strftime("%Y%m%d-%H%M%S") + '.xlsx'
actual_excelwriter = ExcelWriter(filename)
new_sheet = actual_excelwriter.create_sheet('Consolidado')

new_sheet["A1"] = "Semana"
new_sheet["B1"] = "Año"
new_sheet["C1"] = "Flujo"
new_sheet["D1"] = "Tipo"
new_sheet["E1"] = "Rubro"
new_sheet["F1"] = "Fecha"
new_sheet["G1"] = "Cuenta"
new_sheet["H1"] = "Detalle"
new_sheet["I1"] = "Ingreso"
new_sheet["J1"] = "Egreso"

for index, flow in enumerate(actual_flows):
    row_num = str(index + 2)

    new_sheet["A" + row_num] = flow['week']
    new_sheet["B" + row_num] = flow['year']
    new_sheet["C" + row_num] = flow['flow']
    new_sheet["D" + row_num] = flow['type']
    new_sheet["E" + row_num] = flow['category']
    new_sheet["F" + row_num] = flow['date']
    new_sheet["G" + row_num] = flow['account']
    new_sheet["H" + row_num] = flow['details']
    new_sheet["I" + row_num] = flow['income']
    new_sheet["J" + row_num] = flow['expense']

actual_excelwriter.save()

# crear excel de proyecciones
projected_flow_filename = 'outputs/proyecciones_' + time.strftime("%Y%m%d-%H%M%S") + '.xlsx'
projected_excelwriter = ExcelWriter(projected_flow_filename)
projected_sheet = projected_excelwriter.create_sheet('Proyectado')

projected_sheet["A1"] = "Semana"
projected_sheet["B1"] = "Año"
projected_sheet["C1"] = "Flujo"
projected_sheet["D1"] = "Tipo"
projected_sheet["E1"] = "Rubro"
projected_sheet["F1"] = "Fecha"
projected_sheet["G1"] = "Cuenta"
projected_sheet["H1"] = "Detalle"
projected_sheet["I1"] = "Ingreso"
projected_sheet["J1"] = "Egreso"

for index, projected_flow in enumerate(projected_flows):
    row_num = str(index + 2)

    projected_sheet["A" + row_num] = flow['week']
    projected_sheet["B" + row_num] = flow['year']
    projected_sheet["C" + row_num] = flow['flow']
    projected_sheet["D" + row_num] = flow['type']
    projected_sheet["E" + row_num] = flow['category']
    projected_sheet["F" + row_num] = flow['date']
    projected_sheet["G" + row_num] = flow['account']
    projected_sheet["H" + row_num] = flow['details']
    projected_sheet["I" + row_num] = flow['income']
    projected_sheet["J" + row_num] = flow['expense']

projected_excelwriter.save()

# crear excel de categorías faltantes

missing_categories_filename = 'outputs/rubros_faltantes_' + time.strftime("%Y%m%d-%H%M%S") + '.xlsx'
missing_categories_excelwriter = ExcelWriter(missing_categories_filename)
new_categories_sheet = missing_categories_excelwriter.create_sheet('Rubros Faltantes')

new_categories_sheet["A1"] = "CuentaDetalle"
new_categories_sheet["B1"] = "Rubro"
new_categories_sheet["C1"] = "Cuenta"
new_categories_sheet["D1"] = "Detalle"

for index, missing_category_str in enumerate(new_categories):
    row_num = str(index + 2)
    missing_category = json.loads(missing_category_str)

    new_categories_sheet["A" + row_num] = missing_category['account'] + '-' + missing_category['details']
    new_categories_sheet["C" + row_num] = missing_category['account']
    new_categories_sheet["D" + row_num] = missing_category['details']

missing_categories_excelwriter.save()
