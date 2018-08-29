from felsim import translators as translators
from felsim.constants import ACTUAL_FLOW, CHEQUES_TYPE, ESTIMATED_FLOW, INFLEXIBLE


def extract_checks(actual_flows, categories_reader, cheques_sheet, projected_date, projected_flows):
    for rowNum in range(3, cheques_sheet.max_row + 1):  # skip the first row
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

        if date_value and date_value > projected_date:
            check['flow'] = ESTIMATED_FLOW
            check['flexibility'] = INFLEXIBLE
            check['projected_date'] = projected_date.strftime("%d/%m/%Y")
            projected_flows.append(check)
        else:
            actual_flows.append(check)