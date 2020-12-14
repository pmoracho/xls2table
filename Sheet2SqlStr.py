class Sheet2SqlStr(object):

    """Sheet2SqlStr: Clase para transformaciones de un Excel en sentencias SQL."""

    def __init__(self, book, sheet, outputtable, sheet_num, hasheader):
        """__init__."""

        self._book            = book
        self._sheet         = sheet
        self._sql_create    = ''
        self._sql_inserthdr    = ''

        self.max_col         = -1
        self.max_row         = -1
        self.outputtable    = outputtable
        self.sheet_num        = sheet_num
        self.hasheader        = hasheader

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

        if self.outputtable.startswith('#'):
            SQLC = SQLC + "BEGIN TRY\n"
            SQLC = SQLC + "    DROP TABLE " + self.outputtable + "\n"
            SQLC = SQLC + "END TRY\n"
            SQLC = SQLC + "BEGIN CATCH\n"
            SQLC = SQLC + "END CATCH\n\n"

        SQLC = SQLC + "CREATE TABLE {0} (\n".format(self.outputtable)
        SQLC = SQLC + "            ID        INT    IDENTITY,\n"

        SQLI = SQLI + "INSERT INTO {0} ( ".format(self.outputtable)

        if self.hasheader:
            for c in range(0, self.max_col):
                campo = "".join(x for x in str(self._sheet.cell(0, c).value) if x.isalnum())
                SQLC = SQLC + "            {0}        VARCHAR(255),\n".format(campo)
                SQLI = SQLI + "{0}, ".format(campo)
            pass

        else:
            for c in range(0, self.max_col):
                SQLC = SQLC + "            Campo_{0}        VARCHAR(255),\n".format(c+1)
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
            SQLI = self._sql_inserthdr + "    VALUES ("
            for c in range(0, self.max_col):
                cell = self._sheet.cell(r, c)
                # print(cell)
                if cell.value:
                    SQLI = SQLI + "'{0}',".format(self._get_celldata(cell))[:255]
                else:
                    SQLI = SQLI + "NULL,"

            SQLI = SQLI[:-1] + " )\n"
            yield SQLI