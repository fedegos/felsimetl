import time
import sys
from datetime import datetime
import argparse
import pkg_resources

import os

import felsim.categoriesreader as categoriesreader
import felsim.exceladapter as exceladapter

import felsim.excelbuilder as excelbuilder
from felsim.builders.buildmissingcategories import build_missing_categories
from felsim.config.getactualflowstableconfig import get_actual_flows_table_config
from felsim.config.getprojectedflowstableconfig import get_projected_flows_table_config

from felsim.constants import *
from felsim.extractors.extractchecks import extract_checks
from felsim.extractors.extractcash import extract_cash
from felsim.extractors.extractcredicoop import extract_credicoop
from felsim.extractors.extractcurrentaccounts import extract_currentaccounts
from felsim.extractors.extractestimations import extract_estimations


def valid_date(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d")
    except ValueError:
        msg = "No es una fecha válida: '{0}'.".format(s)
        raise argparse.ArgumentTypeError(msg)

def get_parser():
    parser = argparse.ArgumentParser(description='Script para procesar planillas de Felsim')
    parser.add_argument("--version", "-v", help="Muestra la versión.", action="store_true")
    parser.add_argument("--inputs", "-i", help="Permite definir la carpeta de inputs. Si no se indica, el programa busca la carpeta inputs.", default="inputs")
    parser.add_argument("--outputs", "-o", help="Permite definir la carpeta de outputs. Si no se indica, el programa busca la carpeta outputs.", default="outputs")
    parser.add_argument(
        "--date", "-d",
        help="Define la fecha de proyección. Formato: AAAA-MM-DD.",
        required=True,
        type=valid_date
    )
    parser.add_argument("--noprojections", "-np", help="Genera los archivos sin tomar en cuenta las estimaciones.", action="store_true")

    return parser


def main(args=None):
    """The main routine."""
    if args is None:
        args = sys.argv[1:]

    cwd = os.getcwd().replace('\\', '/')

    parser = get_parser()
    params = parser.parse_args(args)

    if params.version:
        print(pkg_resources.require("felsimetl")[0].version)
        return


    inputs_path = "%s/%s/" % (cwd, params.inputs)
    outputs_path = "%s/%s/" % (cwd, params.outputs)
    projected_date = params.date

    print("Fecha de proyección: ", projected_date.strftime("%Y-%m-%d"))

    actual_filename = inputs_path + 'PLANILLA CONTABILIDAD ACTIVA 2018 GSM.xlsx'
    current_accounts_filename = inputs_path + 'CUENTAS CORRIENTES.xlsx'
    categories_filename = inputs_path + 'RUBROS.xlsx'

    include_projections= not params.noprojections

    if not os.path.isdir(inputs_path):
        print("La carpeta %s no existe." % inputs_path)
        return

    if not os.path.isdir(outputs_path):
        print("La carpeta %s no existe." % outputs_path)
        return

    if not os.path.isfile(actual_filename):
        print("El archivo %s no existe." % actual_filename)
        return

    if not os.path.isfile(current_accounts_filename):
        print("El archivo %s no existe." % current_accounts_filename)
        return

    if not os.path.isfile(categories_filename):
        print("El archivo %s no existe." % categories_filename)
        return

    if not include_projections:
        print("No se tomarán en cuenta las estimaciones.")

    categories_reader = categoriesreader.CategoriesReader(categories_filename)
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

    extract_checks(actual_flows, categories_reader, cheques_sheet, projected_date, projected_flows)

    extract_cash(actual_flows, caja_sheet, categories_reader, new_categories)

    extract_credicoop(actual_flows, categories_reader, credicoop_sheet, new_categories, projected_date, projected_flows)

    # generar de hoja planificada (negociable - antes en credicoop abajo de los renglones normales)
    # FLEXIBLE

    if include_projections:
        extract_estimations(categories_reader, estimacion_sheet, new_categories, projected_date, projected_flows)

    # CUENTA CORRIENTE

    extract_currentaccounts(categories_reader, cuentas_corrientes_sheet, new_categories, projected_date,
                            projected_flows)

    # crear excel nuevo
    filename = outputs_path + 'consolidado_real_' + time.strftime("%Y%m%d-%H%M%S") + '.xlsx'
    actual_excelwriter = exceladapter.ExcelWriter(filename)
    new_sheet = actual_excelwriter.create_sheet('Consolidado')

    actual_flow_builder = excelbuilder.BasicBuilder(new_sheet, actual_flows)
    actual_flow_builder.map_configuration(get_actual_flows_table_config())

    actual_flow_builder.build()
    actual_excelwriter.save()

    # crear excel de proyecciones
    projected_flow_filename = outputs_path + 'proyecciones_' + time.strftime("%Y%m%d-%H%M%S") + '.xlsx'
    projected_excelwriter = exceladapter.ExcelWriter(projected_flow_filename)
    projected_sheet = projected_excelwriter.create_sheet('Proyectado')

    projected_flow_builder = excelbuilder.BasicBuilder(projected_sheet, projected_flows)
    projected_flow_builder.map_configuration(get_projected_flows_table_config())

    projected_flow_builder.build()
    projected_excelwriter.save()

    # crear excel de categorías faltantes

    build_missing_categories(new_categories, outputs_path)

    print("Los archivos fueron generados con éxito en la siguiente ubicación:", outputs_path.replace("/", "\\"))


if __name__ == "__main__":
    main()