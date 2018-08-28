import unittest
import subprocess

class xls2tableTest(unittest.TestCase):

	def test_xls2table_sample_exec(self):
		"""test_xls2table_sample_exec: Ejecución de xls2table contra un planilla de referencia Rendimiento.Ejemplo.Importar.xls"""

		reference = """
					--------------------------------------------------------------------------------------------------------
					-- File         : test.xlsx
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

					INSERT INTO prueba ( Campo_1, Campo_2, Campo_3 )        VALUES ('Col1','Col2','Col3' )
					INSERT INTO prueba ( Campo_1, Campo_2, Campo_3 )        VALUES ('AA','1','03-05-2018' )
					INSERT INTO prueba ( Campo_1, Campo_2, Campo_3 )        VALUES ('BB','3','04-05-2018' )
					INSERT INTO prueba ( Campo_1, Campo_2, Campo_3 )        VALUES ('CC','5','05-05-2018' )

					COMMIT TRANSACTION

					--------------------------------------------------------------------------------------------------------
					-- End Script.
					--------------------------------------------------------------------------------------------------------
					"""

		output 		= subprocess.check_output('python xls2table.py test.xlsx prueba "none" -s').decode("utf-8") 
		output 		= ' '.join(output.split())
		reference	= ' '.join(reference.split())
		self.assertEqual(output, reference)

if __name__ == '__main__':
	unittest.main()
