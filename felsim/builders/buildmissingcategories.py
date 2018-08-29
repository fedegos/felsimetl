import json

import time

from felsim import exceladapter as exceladapter, excelbuilder as excelbuilder


def from_json_mapper(item_str, key):
    return json.loads(item_str)[key]

def account_details_maper(item_str, _):
    item = json.loads(item_str)
    result = ("%s-%s" % (item['account'], item['details']))
    return result


def build_missing_categories(new_categories, outputs_path):
    missing_categories_filename = outputs_path + 'rubros_faltantes_' + time.strftime("%Y%m%d-%H%M%S") + '.xlsx'
    missing_categories_excelwriter = exceladapter.ExcelWriter(missing_categories_filename)
    new_categories_sheet = missing_categories_excelwriter.create_sheet('Rubros Faltantes')
    missing_categories_builder = excelbuilder.BasicBuilder(new_categories_sheet, new_categories)
    missing_categories_builder.add_header("A", "CuentaDetalle")
    missing_categories_builder.add_header("B", "Rubro")
    missing_categories_builder.add_header("C", "Cuenta")
    missing_categories_builder.add_header("D", "Detalle")

    missing_categories_builder.map_column("A", "account", account_details_maper)
    missing_categories_builder.map_column("C", "account", from_json_mapper)
    missing_categories_builder.map_column("D", "details", from_json_mapper)
    missing_categories_builder.build()
    missing_categories_excelwriter.save()