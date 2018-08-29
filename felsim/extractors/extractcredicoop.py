from felsim import tableextraction as tableextraction, translators as translators
from felsim.constants import CARGAS_SOCIALES, SUELDOS, BANKS, PRESTAMOS_BANCARIOS, ANNULED, T_E_CUENTAS_PROPIAS, \
    MORATORIA_AFIP, MORATORIAS, NEW_CATEGORY, ACTUAL_FLOW, CREDICOOP_TYPE, ESTIMATED_FLOW, INFLEXIBLE


def extract_credicoop(actual_flows, categories_reader, credicoop_sheet, new_categories, projected_date,
                      projected_flows):
    credicoop_unpacker = tableextraction.TableUnpacker(credicoop_sheet)
    for rowNum in range(3, credicoop_sheet.max_row + 1):
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
        if date_value and date_value > projected_date:
            cash_flow['flow'] = ESTIMATED_FLOW
            cash_flow['flexibility'] = INFLEXIBLE
            cash_flow['projected_date'] = projected_date.strftime("%d/%m/%Y")
            projected_flows.append(cash_flow)
        else:
            actual_flows.append(cash_flow)