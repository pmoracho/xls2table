import csv

class Csv2SqlStr(object):

    """Csv2SqlStr: Clase para transformaciones de un CSV en sentencias SQL."""

    def __init__(self, file, outputtable, hasheader, sep=";"):
        """__init__."""

        self._file            = file
        self._sql_create    = ''
        self._sql_inserthdr    = ''
        self._rows = []

        self.max_col         = -1
        self.max_row         = -1
        self.outputtable    = outputtable
        self.hasheader        = hasheader

        with open(self._file) as f:
            csv_reader_object = csv.reader(f, delimiter=sep)
            self._rows = [line for line in csv_reader_object]

        self._load_sheet_limits()
        self._create_header_stmts()

    def _load_sheet_limits(self):
        """_load_sheet_limits: Determina la fila y columna máxima de la hoja."""
        self.max_col = 0
        self.max_row = len(self._rows)
        for row in self._rows:
            cols = len(row)
            if cols > self.max_col:
                self.max_col = cols


    def _create_header_stmts(self):
        """_create_header_stmts: Crea los string de cabecera de las sentencias de insert y creación."""

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
            for row in self._rows[0]:
                campo = "".join(str(x).strip() for x in row)
                SQLC = SQLC + "            {0}         VARCHAR(255),\n".format(campo)
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
        for row in self._rows:

            SQLI = self._sql_inserthdr + "  VALUES ("
            for c in range(0, self.max_col):
                cell = row[c]
                # print(cell)
                if cell:
                    SQLI = SQLI + "'{0}',".format(cell)[:255].strip()
                else:
                    SQLI = SQLI + "NULL,"

            SQLI = SQLI[:-1] + ")\n"
            yield SQLI


def test():

    cvsfile = "prueba.csv"

    S2Sql     = Csv2SqlStr(cvsfile, "tabla", False)

    SQL_start     = "\nBEGIN TRANSACTION\n\n"
    SQL_start     = SQL_start + S2Sql.get_create_sql()
    SQL_start     = SQL_start + "\n"

    SQL_rows    = []
    SQL_rows     = [isql for isql in S2Sql.get_insert_stmts()]

    SQL_end     = "\nCOMMIT TRANSACTION\n"

    print(SQL_start + "".join(SQL_rows) + SQL_end)

if __name__ == "__main__":
    test()