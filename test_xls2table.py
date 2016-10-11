import unittest
import subprocess

class xls2tableTest(unittest.TestCase):

	def test_xls2table_sample_exec(self):
		"""test_xls2table_sample_exec: Ejecución de xls2table contra un planilla de referencia Rendimiento.Ejemplo.Importar.xls"""

		reference = """
					--------------------------------------------------------------------------------------------------------
					-- File         : Rendimiento.Ejemplo.Importar.xls
					-- Output table : prueba
					-- Dsn          : none
					--------------------------------------------------------------------------------------------------------

					BEGIN TRANSACTION

					BEGIN TRY
							DROP TABLE prueba
					END TRY
					BEGIN CATCH
					END CATCH

					CREATE TABLE prueba (
											ID              INT     IDENTITY,
											Campo_1         VARCHAR(255),
											Campo_2         VARCHAR(255),
											Campo_3         VARCHAR(255)
					)

					INSERT INTO prueba ( Campo_1, Campo_2, Campo_3 )
							VALUES ('AMOF','12-12-2011','1000' )
					INSERT INTO prueba ( Campo_1, Campo_2, Campo_3 )
							VALUES ('GG','01-01-2011','550.6' )
					INSERT INTO prueba ( Campo_1, Campo_2, Campo_3 )
							VALUES ('FMO','01-01-2011','563' )
					INSERT INTO prueba ( Campo_1, Campo_2, Campo_3 )
							VALUES ('ES','01-01-2012','89' )
					INSERT INTO prueba ( Campo_1, Campo_2, Campo_3 )
							VALUES ('MSMAR','05-10-2012','154' )
					INSERT INTO prueba ( Campo_1, Campo_2, Campo_3 )
							VALUES ('GPLP','05-10-2007','632' )

					COMMIT TRANSACTION

					--------------------------------------------------------------------------------------------------------
					-- End Script.
					--------------------------------------------------------------------------------------------------------
					"""

		output 		= subprocess.check_output('python xls2table.py Rendimiento.Ejemplo.Importar.xls prueba "none"  -s').decode("utf-8") 
		output 		= ' '.join(output.split())
		reference	= ' '.join(reference.split())
		self.assertEqual(output, reference)

if __name__ == '__main__':
	unittest.main()
