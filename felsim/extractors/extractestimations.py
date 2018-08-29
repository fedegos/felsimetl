from felsim import tableextraction as tableextraction, translators as translators
from felsim.constants import CARGAS_SOCIALES, SUELDOS, BANKS, PRESTAMOS_BANCARIOS, ANNULED, T_E_CUENTAS_PROPIAS, \
    MORATORIA_AFIP, MORATORIAS, NEW_CATEGORY, ESTIMATED_FLOW, FLEXIBLE, ESTIMADO_TYPE


def extract_estimations(categories_reader, estimacion_sheet, new_categories, projected_date, projected_flows):
    estimacion_unpacker = tableextraction.TableUnpacker(estimacion_sheet)
    for row_num in range(2, estimacion_sheet.max_row + 1):
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

        cash_flow['projected_date'] = projected_date.strftime("%d/%m/%Y")
        projected_flows.append(cash_flow)