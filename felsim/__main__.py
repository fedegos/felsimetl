import openpyxl
import time
import json
import sys
from datetime import datetime, timedelta
# from felsim.constants import *
from constants import *

from functools import reduce

import os

import tableextraction
import categoriesreader
import exceladapter
import translators

import excelbuilder

def main(args=None):
    # abrir
    """The main routine."""
    if args is None:
        args = sys.argv[1:]

    cwd = os.getcwd().replace('\\', '/')
    # cwd = os.path.abspath(os.path.dirname(__file__))
    # cwd = 'C:/Users/Federico/PycharmProjects/felsim'
    print(cwd)

    actual_filename = cwd + '/inputs/PLANILLA CONTABILIDAD ACTIVA 2018 GSM.xlsx'
    current_accounts_filename = cwd + '/inputs/CUENTAS CORRIENTES.xlsx'

    categories_reader = categoriesreader.CategoriesReader(cwd + '/inputs/RUBROS.xlsx')
    actual_excelreader = exceladapter.excelreader.ExcelReader(actual_filename)
    current_accounts_excelreader = exceladapter.ExcelReader(current_accounts_filename)

    cheques_sheet = actual_excelreader.get_sheet(CHEQUES_SHEET_NAME)
    caja_sheet = actual_excelreader.get_sheet(CAJA_SHEET_NAME)
    credicoop_sheet = actual_excelreader.get_sheet(CREDICOOP_SHEET_NAME)
    estimacion_sheet = actual_excelreader.get_sheet(ESTIMACION_SHEET_NAME)

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
        check['expense'] = translators.to_google_num(providers_amount)

        if date_value and date_value > datetime.now():
            check['flow'] = ESTIMATED_FLOW
            check['flexibility'] = INFLEXIBLE
            projected_flows.append(check)
        else:
            actual_flows.append(check)

    table_unpacker = tableextraction.TableUnpacker(caja_sheet)

    for rowNum in range(3, caja_sheet.max_row):
        cash_flow = {}

        row_unpacker = table_unpacker.get_row_unpacker(rowNum)

        account = row_unpacker.get_value_at(3)
        details = row_unpacker.get_value_at(2)
        date_value = row_unpacker.get_value_at(1)

        if not date_value:
            break

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

        cash_flow['date'], cash_flow['week'], cash_flow['year'] = translators.unpack_dates(date_value)

        cash_flow['flow'] = ACTUAL_FLOW
        cash_flow['flexibility'] = ""

        cash_flow['type'] = CAJAS_TYPE  # caja

        cash_flow['details'] = details

        cash_flow['category'] = category
        cash_flow['account'] = account
        cash_flow['income'] = ""
        cash_flow['expense'] = translators.to_google_num(expense) if expense else ""
        cash_flow['income'] = translators.to_google_num(income) if income else ""

        # print(cash_flow)



        actual_flows.append(cash_flow)

    credicoop_unpacker = tableextraction.TableUnpacker(credicoop_sheet)

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

        if not details:
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

        cash_flow['date'], cash_flow['week'], cash_flow['year'] = translators.unpack_dates(date_value)

        cash_flow['flow'] = ACTUAL_FLOW
        cash_flow['flexibility'] = ""
        cash_flow['type'] = CREDICOOP_TYPE
        cash_flow['details'] = details

        cash_flow['category'] = category
        cash_flow['account'] = account
        cash_flow['income'] = ""
        cash_flow['expense'] = translators.to_google_num(expense) if expense else ""
        cash_flow['income'] = translators.to_google_num(income) if income else ""

        # print(cash_flow)
        if date_value and date_value > datetime.now():
            cash_flow['flow'] = ESTIMATED_FLOW
            cash_flow['flexibility'] = INFLEXIBLE
            projected_flows.append(cash_flow)
        else:
            actual_flows.append(cash_flow)

    # generar de hoja planificada (negociable - antes en credicoop abajo de los renglones normales)
    # FLEXIBLE

    estimacion_unpacker = tableextraction.TableUnpacker(estimacion_sheet)

    for row_num in range(2, estimacion_sheet.max_row):
        cash_flow = {}

        row_unpacker = estimacion_unpacker.get_row_unpacker(row_num)

        details = row_unpacker.get_value_at(6)
        account = row_unpacker.get_value_at(7)

        if details == CARGAS_SOCIALES:
            account = SUELDOS

        if details in BANKS:
            account = PRESTAMOS_BANCARIOS

        if details == ANNULED:
            continue

        if not details:
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

        date_value = row_unpacker.get_value_at(4)

        cash_flow['date'], cash_flow['week'], cash_flow['year'] = translators.unpack_dates(date_value)

        cash_flow['flow'] = ESTIMATED_FLOW
        cash_flow['flexibility'] = FLEXIBLE
        cash_flow['type'] = ESTIMADO_TYPE
        cash_flow['details'] = details

        cash_flow['category'] = category
        cash_flow['account'] = account
        cash_flow['expense'] = translators.to_google_num(expense) if expense else ""
        cash_flow['income'] = translators.to_google_num(income) if income else ""

        projected_flows.append(cash_flow)

    # CUENTA CORRIENTE

    sales_by_customer = {}
    weeks_projected_by_customer = {}

    current_date = datetime.now()
    current_date_at_zero = datetime(current_date.year, current_date.month, current_date.day)
    _, current_week, current_year = translators.unpack_dates(current_date)

    def get_relative_weeks(since_date, until_date):
        since_date_at_zero = datetime(since_date.year, since_date.month, since_date.day)
        monday1 = (since_date_at_zero - timedelta(days=since_date_at_zero.weekday()))
        monday2 = (until_date - timedelta(days=until_date.weekday()))
        return (monday2 - monday1).days // 7

    def add_weeks(since_date, number_of_weeks):
        return since_date + timedelta(days=number_of_weeks * 7)

    cuentas_corrientes_unpacker = tableextraction.TableUnpacker(cuentas_corrientes_sheet)

    for row_unpacker in cuentas_corrientes_unpacker.read_rows(2):
        cash_flow = {}

        customer = details = row_unpacker.get_value_at(1)
        date_value = row_unpacker.get_value_at(2)
        income = row_unpacker.get_value_at(3)

        if not date_value:
            continue

        if date_value <= datetime.now():
            continue

        category, account = categories_reader.get_category_from_details(details)

        if category == "" and account == "":
            new_categories.add('{"account": "*** VENTAS A CLASIFICAR ***", "details": "%s"}' % (details))

        cash_flow['date'], cash_flow['week'], cash_flow['year'] = translators.unpack_dates(date_value)

        # year can be greater than 1
        relative_weeks = get_relative_weeks(current_date, date_value)

        if not details in sales_by_customer:
            sales_by_customer[customer] = []

        if not details in weeks_projected_by_customer:
            weeks_projected_by_customer[customer] = 0

        if relative_weeks > weeks_projected_by_customer[customer]:
            weeks_projected_by_customer[customer] = relative_weeks

        sales_by_customer[details].append(income)

        cash_flow['flow'] = ESTIMATED_FLOW
        cash_flow['flexibility'] = INFLEXIBLE
        cash_flow['type'] = ESTIMADO_TYPE
        cash_flow['details'] = details

        cash_flow['category'] = category
        cash_flow['account'] = account
        cash_flow['expense'] = ''
        cash_flow['income'] = translators.to_google_num(income) if income else ""

        projected_flows.append(cash_flow)

    for customer, customer_sales in sales_by_customer.items():
        average_income =  reduce(lambda x, y: x + y, customer_sales)/ len(customer_sales)
        # average_income_2 = reduce(lambda x, y: x + y, customer_sales) / weeks_projected_by_customer[customer]

        project_since = weeks_projected_by_customer[customer] + 1

        for week in range(project_since, 14):
            cash_flow = {}

            projected_date = add_weeks(current_date_at_zero, week)

            category, account = categories_reader.get_category_from_details(customer)

            cash_flow['date'], cash_flow['week'], cash_flow['year'] = translators.unpack_dates(projected_date)

            cash_flow['flow'] = ESTIMATED_FLOW
            cash_flow['flexibility'] = FLEXIBLE
            cash_flow['type'] = ESTIMADO_TYPE
            cash_flow['details'] = customer

            cash_flow['category'] = category
            cash_flow['account'] = account
            cash_flow['expense'] = ''
            cash_flow['income'] = translators.to_google_num(average_income) if average_income else ""

            projected_flows.append(cash_flow)

    # crear excel nuevo
    filename = cwd + '/outputs/consolidado_real_' + time.strftime("%Y%m%d-%H%M%S") + '.xlsx'
    actual_excelwriter = exceladapter.ExcelWriter(filename)
    new_sheet = actual_excelwriter.create_sheet('Consolidado')

    actual_flow_builder = excelbuilder.BasicBuilder(new_sheet, actual_flows)

    actual_flow_builder.add_header("A", "Semana")
    actual_flow_builder.add_header("B", "Año")
    actual_flow_builder.add_header("C", "Flujo")
    actual_flow_builder.add_header("D", "Flexibilidad")
    actual_flow_builder.add_header("E", "Tipo")
    actual_flow_builder.add_header("F", "Rubro")
    actual_flow_builder.add_header("G", "Fecha")
    actual_flow_builder.add_header("H", "Cuenta")
    actual_flow_builder.add_header("I", "Detalle")
    actual_flow_builder.add_header("J", "Ingreso")
    actual_flow_builder.add_header("K", "Egreso")

    actual_flow_builder.map_column("A", "week")
    actual_flow_builder.map_column("B", "year")
    actual_flow_builder.map_column("C", "flow")
    actual_flow_builder.map_column("D", "flexibility")
    actual_flow_builder.map_column("E", "type")
    actual_flow_builder.map_column("F", "category")
    actual_flow_builder.map_column("G", "date")
    actual_flow_builder.map_column("H", "account")
    actual_flow_builder.map_column("I", "details")
    actual_flow_builder.map_column("J", "income")
    actual_flow_builder.map_column("K", "expense")

    actual_flow_builder.build()
    actual_excelwriter.save()

    # crear excel de proyecciones
    projected_flow_filename = cwd + '/outputs/proyecciones_' + time.strftime("%Y%m%d-%H%M%S") + '.xlsx'
    projected_excelwriter = exceladapter.ExcelWriter(projected_flow_filename)
    projected_sheet = projected_excelwriter.create_sheet('Proyectado')

    projected_flow_builder = excelbuilder.BasicBuilder(projected_sheet, projected_flows)

    projected_flow_builder.add_header("A", "Semana")
    projected_flow_builder.add_header("B", "Año")
    projected_flow_builder.add_header("C", "Flujo")
    projected_flow_builder.add_header("D", "Flexibilidad")
    projected_flow_builder.add_header("E", "Tipo")
    projected_flow_builder.add_header("F", "Rubro")
    projected_flow_builder.add_header("G", "Fecha")
    projected_flow_builder.add_header("H", "Cuenta")
    projected_flow_builder.add_header("I", "Detalle")
    projected_flow_builder.add_header("J", "Ingreso")
    projected_flow_builder.add_header("K", "Egreso")

    projected_flow_builder.map_column("A", "week")
    projected_flow_builder.map_column("B", "year")
    projected_flow_builder.map_column("C", "flow")
    projected_flow_builder.map_column("D", "flexibility")
    projected_flow_builder.map_column("E", "type")
    projected_flow_builder.map_column("F", "category")
    projected_flow_builder.map_column("G", "date")
    projected_flow_builder.map_column("H", "account")
    projected_flow_builder.map_column("I", "details")
    projected_flow_builder.map_column("J", "income")
    projected_flow_builder.map_column("K", "expense")

    projected_flow_builder.build()
    projected_excelwriter.save()

    # crear excel de categorías faltantes

    missing_categories_filename = cwd + '/outputs/rubros_faltantes_' + time.strftime("%Y%m%d-%H%M%S") + '.xlsx'
    missing_categories_excelwriter = exceladapter.ExcelWriter(missing_categories_filename)
    new_categories_sheet = missing_categories_excelwriter.create_sheet('Rubros Faltantes')

    missing_categories_builder = excelbuilder.BasicBuilder(new_categories_sheet, new_categories)

    missing_categories_builder.add_header("A", "CuentaDetalle")
    missing_categories_builder.add_header("B", "Rubro")
    missing_categories_builder.add_header("C", "Cuenta")
    missing_categories_builder.add_header("D", "Detalle")

    def from_json_mapper(item_str, key):
        return json.loads(item_str)[key]


    def account_details_maper(item_str, _):
        item = json.loads(item_str)

        result = ("%s-%s" % (item['account'], item['details']))
        return result

    missing_categories_builder.map_column("A", "account", account_details_maper)
    missing_categories_builder.map_column("C", "account", from_json_mapper)
    missing_categories_builder.map_column("D", "details", from_json_mapper)

    missing_categories_builder.build()

    missing_categories_excelwriter.save()

if __name__ == "__main__":
    main()