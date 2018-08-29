from datetime import datetime, timedelta
from functools import reduce

from felsim import translators as translators, tableextraction as tableextraction
from felsim.constants import ESTIMATED_FLOW, INFLEXIBLE, ESTIMADO_TYPE, FLEXIBLE

def get_relative_weeks(since_date, until_date):
    since_date_at_zero = datetime(since_date.year, since_date.month, since_date.day)
    monday1 = (since_date_at_zero - timedelta(days=since_date_at_zero.weekday()))
    monday2 = (until_date - timedelta(days=until_date.weekday()))
    return (monday2 - monday1).days // 7


def add_weeks(since_date, number_of_weeks):
    return since_date + timedelta(days=number_of_weeks * 7)


def extract_currentaccounts(categories_reader, cuentas_corrientes_sheet, new_categories, projected_date,
                            projected_flows):
    sales_by_customer = {}
    weeks_projected_by_customer = {}
    current_date = projected_date
    current_date_at_zero = datetime(current_date.year, current_date.month, current_date.day)
    _, current_week, current_year = translators.unpack_dates(current_date)



    cuentas_corrientes_unpacker = tableextraction.TableUnpacker(cuentas_corrientes_sheet)
    for row_unpacker in cuentas_corrientes_unpacker.read_rows(2):
        cash_flow = {}

        customer = details = row_unpacker.get_value_at(1)
        date_value = row_unpacker.get_value_at(2)
        income = row_unpacker.get_value_at(3)

        if not date_value:
            continue

        if date_value <= projected_date:
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

        cash_flow['projected_date'] = projected_date.strftime("%d/%m/%Y")
        projected_flows.append(cash_flow)

    for customer, customer_sales in sales_by_customer.items():
        average_income = reduce(lambda x, y: x + y, customer_sales) / len(customer_sales)
        # average_income_2 = reduce(lambda x, y: x + y, customer_sales) / weeks_projected_by_customer[customer]

        project_since = weeks_projected_by_customer[customer] + 1

        for week in range(project_since, 14):
            cash_flow = {}

            projected_sales_date = add_weeks(current_date_at_zero, week)

            category, account = categories_reader.get_category_from_details(customer)

            cash_flow['date'], cash_flow['week'], cash_flow['year'] = translators.unpack_dates(projected_sales_date)

            cash_flow['flow'] = ESTIMATED_FLOW
            cash_flow['flexibility'] = FLEXIBLE
            cash_flow['type'] = ESTIMADO_TYPE
            cash_flow['details'] = customer

            cash_flow['category'] = category
            cash_flow['account'] = account
            cash_flow['expense'] = ''
            cash_flow['income'] = translators.to_google_num(average_income) if average_income else ""

            cash_flow['projected_date'] = projected_date.strftime("%d/%m/%Y")
            projected_flows.append(cash_flow)