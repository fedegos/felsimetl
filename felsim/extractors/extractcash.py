from felsim import tableextraction as tableextraction, translators as translators
from felsim.constants import CARGAS_SOCIALES, SUELDOS, BANKS, PRESTAMOS_BANCARIOS, T_E_CUENTAS_PROPIAS, \
    FELSIM_CREDICOOP, NEW_CATEGORY, ACTUAL_FLOW, CAJAS_TYPE


def extract_cash(actual_flows, caja_sheet, categories_reader, new_categories):
    table_unpacker = tableextraction.TableUnpacker(caja_sheet)
    for rowNum in range(3, caja_sheet.max_row + 1):
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

        actual_flows.append(cash_flow)