# -*- coding: utf-8 -*-

# Copyright (c) 2014 Patricio Moracho <pmoracho@gmail.com>
#
# xls2table
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of version 3 of the GNU General Public License
# as published by the Free Software Foundation. A copy of this license should
# be included in the file GPL-3.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Library General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.

__author__        = "Patricio Moracho <pmoracho@gmail.com>"
__appname__        = "xls2table"
__appdesc__        = "Importa datos de Excel a una tabla SQL"
__license__        = 'GPL v3'
__copyright__    = "(c) 2016, %s" % (__author__)
__version__     = "1.0.1"
__date__        = "2016/01/14"

"""
###############################################################################
# Imports
###############################################################################
"""
try:
    import sys
    # import os
    import datetime
    import gettext

    def my_gettext(s):
        """my_gettext: Traducir algunas cadenas de argparse."""
        current_dict = {
            'usage: ': 'uso: ',
            'optional arguments': 'argumentos opcionales',
            'show this help message and exit': 'mostrar esta ayuda y salir',
            'positional arguments': 'argumentos posicionales',
            'the following arguments are required: %s': 'los siguientes argumentos son requeridos: %s'
        }

        if s in current_dict:
            return current_dict[s]
        return s

    gettext.gettext = my_gettext

    import logging
    import argparse

    """
    Librerias adicionales
    """
    import xlrd
    # from xlrd.sheet import ctype_text
    import pypyodbc
    from Csv2SqlStr import Csv2SqlStr
    from Sheet2SqlStr import Sheet2SqlStr

except ImportError as err:
    modulename = err.args[0].partition("'")[-1].rpartition("'")[0]
    print("No fue posible importar el modulo: %s" % modulename)
    sys.exit(-1)


def init_argparse():
    """init_argparse: Inicializar parametros del programa."""
    cmdparser = argparse.ArgumentParser(
                                    prog=__appname__,
                                    description="%s\n%s\n" % (__appdesc__, __copyright__),
                                    epilog="",
                                    add_help=True,
                                    formatter_class=lambda prog: argparse.HelpFormatter(prog, max_help_position=40)
    )

    cmdparser.add_argument('inputfile'                , type=str,                  action="store"                             , help="Archivo Excel de entrada")
    cmdparser.add_argument('outputtable'            , type=str,                  action="store"                             , help="Tabla SQL donde se insertarán las filas de la planilla")
    cmdparser.add_argument('dsn'                    , type=str,                  action="store"                             , help="Cadena de conexión")

    cmdparser.add_argument('-v', '--version'         , action='version', version=__version__                                , help='Mostrar el número de versión y salir')
    cmdparser.add_argument('-l', '--log'            , type=str,                  action="store", dest="log"                , help="Nivel de log", metavar="<level>", default="none")
    cmdparser.add_argument('-n', '--sheetnum'        , type=int,                 action="store", dest="sheet_num"        , help="Número de solapa a leer (la primera es 0)", metavar="<numero>", default="0")
    cmdparser.add_argument('-c', '--hasheader'        ,                              action="store_true", dest="hasheader"    , help="La planilla tiene en la primer fila el nombre de los campos?")
    cmdparser.add_argument('-t', '--nativetypes'    ,                              action="store_true", dest="nativetypes"    , help="Intentar respetar el tipo de dato de cada columna, sino los campo por defecto se generan como VARCHAR(255)")
    cmdparser.add_argument('-s', '--showonly'        ,                              action="store_true", dest="showonly"    , help="Solo muestra el Script a ejecutar")

    return cmdparser




def chunks(lineas, maxlen):

    batch = ""
    maxlinea = max([len(l) for l in lineas])
    for l in lineas:
        batch = batch + l
        if len(batch) + maxlinea >= maxlen:
          yield batch
          batch = ""

    if batch != "":
        yield batch

def rows(lineas, maxlineas):

    batch = ""
    i = 0
    for l in lineas:
        batch = batch + l
        i = i + 1
        if i >= maxlineas:
          yield batch
          batch = ""
          i = 0

    if batch != "":
        yield batch

def procxls(inputfile, outputtable, dsn, sheet_num, hasheader, showonly):
    # """procxls: Proceso principal de importación del xls"""

    logging.info("Procesando %s" % inputfile)

    # if not outputtable.startswith('#'):
    #     raise ValueError('Solo se permite importar a una tabla temporal')

    if inputfile[-3:].lower() == "csv":
        S2Sql     = Csv2SqlStr(inputfile, outputtable, hasheader)
    else:

        book     = xlrd.open_workbook(inputfile)
        sheet    = book.sheet_by_index(sheet_num)

        S2Sql     = Sheet2SqlStr(book, sheet, outputtable, sheet_num, hasheader)

    SQL_start     = "\nBEGIN TRANSACTION\n\n"
    SQL_start     = SQL_start + S2Sql.get_create_sql()
    SQL_start     = SQL_start + "\n"

    SQL_rows    = []
    SQL_rows     = [isql for isql in S2Sql.get_insert_stmts()]

    SQL_end     = "\nCOMMIT TRANSACTION\n"

    logging.debug(SQL_start + "".join(SQL_rows) + SQL_end)

    logging.info("Ejecutando inserción en la tabla %s" % outputtable)

    if not showonly:
        try:
            conn     = pypyodbc.connect(dsn)
            cur        = conn.cursor()

            cur.execute(SQL_start)

            # Armamos lotes de nomas de 100 Kb
            for batch in rows(SQL_rows, 100):
                cur.execute("".join(batch))

            cur.execute(SQL_end)

            cur.commit()
            conn.close()

        except Exception as e:
            logging.error("%s error: %s" % (__appname__, str(e)))
    else:
        print("")
        print("--------------------------------------------------------------------------------------------------------")
        print("-- File         : {0}".format(inputfile))
        print("-- Output table : {0}".format(outputtable))
        print("-- Dsn          : {0}".format(dsn))
        print("--------------------------------------------------------------------------------------------------------")
        print(SQL_start + "".join(SQL_rows) + SQL_end)
        print("--------------------------------------------------------------------------------------------------------")
        print("-- End Script.")
        print("--------------------------------------------------------------------------------------------------------")


"""
##################################################################################################################################################
# Main program
##################################################################################################################################################
"""
if __name__ == "__main__":

    cmdparser = init_argparse()
    try:
        args = cmdparser.parse_args()
    except IOError as msg:
        args.error(str(msg))

    try:
        log_level = getattr(logging, args.log.upper(), None)
        if not isinstance(log_level, int):
            log_level = 51  # No log

        logging.basicConfig(level=log_level, format='%(asctime)s:%(levelname)s:%(message)s')

        logging.info('iniciando proceso')
        procxls(args.inputfile, args.outputtable, args.dsn, args.sheet_num, args.hasheader, args.showonly)
        logging.info('Fin del proceso')

    except Exception as e:
        logging.error("%s error: %s" % (__appname__, str(e)))

    sys.exit(0)
