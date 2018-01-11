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

__author__		= "Patricio Moracho <pmoracho@gmail.com>"
__appname__		= "xls2table"
__appdesc__		= "Importa datos de Excel a una tabla SQL"
__license__		= 'GPL v3'
__copyright__	= "(c) 2016, %s" % (__author__)
__version__ 	= "1.0.1"
__date__		= "2016/01/14"

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

	cmdparser.add_argument('inputfile'				, type=str,  				action="store"	 						, help="Archivo Excel de entrada")
	cmdparser.add_argument('outputtable'			, type=str,  				action="store"	 						, help="Tabla SQL donde se insertarán las filas de la planilla")
	cmdparser.add_argument('dsn'					, type=str,  				action="store"	 						, help="Cadena de conexión")

	cmdparser.add_argument('-v', '--version'     	, action='version', version=__version__								, help='Mostrar el número de versión y salir')
	cmdparser.add_argument('-l', '--log'			, type=str,  				action="store", dest="log"				, help="Nivel de log", metavar="<level>", default="none")
	cmdparser.add_argument('-n', '--sheetnum'		, type=int, 				action="store", dest="sheet_num"		, help="Número de solapa a leer (la primera es 0)", metavar="<numero>", default="0")
	cmdparser.add_argument('-c', '--hasheader'		, 			 				action="store_true", dest="hasheader"	, help="La planilla tiene en la primer fila el nombre de los campos?")
	cmdparser.add_argument('-t', '--nativetypes'	, 			 				action="store_true", dest="nativetypes"	, help="Intentar respetar el tipo de dato de cada columna, sino los campo por defecto se generan como VARCHAR(255)")
	cmdparser.add_argument('-s', '--showonly'		, 			 				action="store_true", dest="showonly"	, help="Solo muestra el Script a ejecutar")

	return cmdparser


class Sheet2SqlStr(object):

	"""Sheet2SqlStr: Clase para transformaciones de un Excel en sentencias SQL."""

	def __init__(self, book, sheet, outputtable, sheet_num, hasheader):
		"""__init__."""

		self._book			= book
		self._sheet 		= sheet
		self._sql_create	= ''
		self._sql_inserthdr	= ''

		self.max_col 		= -1
		self.max_row 		= -1
		self.outputtable	= outputtable
		self.sheet_num		= sheet_num
		self.hasheader		= hasheader

		self._load_sheet_limits()
		self._create_header_stmts()

	def _load_sheet_limits(self):
		"""_load_sheet_limits: Determina la fila y columna máxima de la hoja."""
		self.max_col = 0
		self.max_row = self._sheet.nrows
		for row_idx in range(0, self._sheet.nrows):
			cols = len(self._sheet.row(row_idx))
			if cols > self.max_col:
				self.max_col = cols

	def _get_celldata(self, cell):
		"""_get_celldata: Retorna los datos como strings"""
		if cell.ctype == xlrd.XL_CELL_DATE:
			# Returns a tuple.
			dt_tuple = xlrd.xldate_as_tuple(cell.value, self._book.datemode)
			# Create datetime object from this tuple.
			dt = (datetime.datetime(
							dt_tuple[0], dt_tuple[1], dt_tuple[2],
							dt_tuple[3], dt_tuple[4], dt_tuple[5]
							))
			return dt.strftime("%d-%m-%Y")
		elif cell.ctype == xlrd.XL_CELL_NUMBER:
			return str(int(cell.value)) if int(cell.value) == cell.value else str(cell.value)
		elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
			return str(cell.value).lower()
		else:
			return str(cell.value).replace("'", "''")
		return "xx"

	def _create_header_stmts(self):
		"""_create_header_stmts: Crea los string de cabecera de las sentecias de insert y creación."""

		SQLC = ""
		SQLI = ""

		SQLC = SQLC + "BEGIN TRY\n"
		SQLC = SQLC + "	DROP TABLE " + self.outputtable + "\n"
		SQLC = SQLC + "END TRY\n"
		SQLC = SQLC + "BEGIN CATCH\n"
		SQLC = SQLC + "END CATCH\n\n"

		SQLC = SQLC + "CREATE TABLE {0} (\n".format(self.outputtable)
		SQLC = SQLC + "			ID		INT	IDENTITY,\n"

		SQLI = SQLI + "INSERT INTO {0} ( ".format(self.outputtable)

		if self.hasheader:
			for c in range(0, self.max_col):
				campo = "".join(x for x in str(self._sheet.cell(0, c).value) if x.isalnum())
				SQLC = SQLC + "			{0}		VARCHAR(255),\n".format(campo)
				SQLI = SQLI + "{0}, ".format(campo)
			pass

		else:
			for c in range(0, self.max_col):
				SQLC = SQLC + "			Campo_{0}		VARCHAR(255),\n".format(c+1)
				SQLI = SQLI + "Campo_{0}, ".format(c+1)

			SQLC = SQLC[:-2] + "\n)\n"
			SQLI = SQLI[:-2] + " ) "

		self._sql_create = SQLC
		self._sql_inserthdr = SQLI

	def get_create_sql(self):
		"""get_create_sql: devuelve el SQL de creación de tabla."""
		return self._sql_create

	def get_insert_stmts(self):
		"""get_insert_stmts: devuelve el SQL de inser de una fila."""
		for r in range(1 if self.hasheader else 0, self.max_row):
			SQLI = self._sql_inserthdr + "	VALUES ("
			for c in range(0, self.max_col):
				cell = self._sheet.cell(r, c)
				# print(cell)
				if cell.value:
					SQLI = SQLI + "'{0}',".format(self._get_celldata(cell))[:255]
				else:
					SQLI = SQLI + "NULL,"

			SQLI = SQLI[:-1] + " )\n"
			yield SQLI

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

def procxls(inputfile, outputtable, dsn, sheet_num, hasheader, showonly):
	# """procxls: Proceso principal de importación del xls"""

	logging.info("Procesando %s" % inputfile)

	book 	= xlrd.open_workbook(inputfile)
	sheet	= book.sheet_by_index(sheet_num)
	S2Sql 	= Sheet2SqlStr(book, sheet, outputtable, sheet_num, hasheader)

	SQL_start 	= "\nBEGIN TRANSACTION\n\n"
	SQL_start 	= SQL_start + S2Sql.get_create_sql()
	SQL_start 	= SQL_start + "\n"

	SQL_rows	= []
	SQL_rows 	= [isql for isql in S2Sql.get_insert_stmts()]

	SQL_end 	= "\nCOMMIT TRANSACTION\n"

	logging.debug(SQL_start + "".join(SQL_rows) + SQL_end)

	logging.info("Ejecutando inserción en la tabla %s" % outputtable)

	if not showonly:
		try:
			conn 	= pypyodbc.connect(dsn)
			cur		= conn.cursor()

			cur.execute(SQL_start)

			# Armamos lotes de nomas de 100 Kb
			for batch in chunks(SQL_rows, 100000):
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
