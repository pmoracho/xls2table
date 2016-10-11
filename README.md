Xls2table
===============
Herramienta para convertir una planilla Excel (xls, xlsx) en un tabla SQL. La lectura del formato es nativa y no requiere de Microsoft Excel instalado.

Pasos previos Desarrollo
========================

Para desarrollo de la herramienta es necesario, además de contar con el entorno de desarrollo python mencionado [aquí](../README.md), tener en cuenta la siguiente información:

* Crear el entorno de desarrollo
	* Crear el entorno virtual, de esta manera aislamos las librerías que necesitaremos sin "ensuciar" el entorn Python base: `virtualenv ../venvs/xls2table`
	* Activar el entorno, antes que nada hay que activar el entorno, para que los paths a Python apunten a las nuevas carpetas:  `source  ../venvs/xls2table/Scripts/activate`
	* Instalar librerías adicionales. 
		* [xlrd](https://github.com/python-excel/xlrd) para la lectura del formato excel: `pip install xlrd`
		* [pypyodbc](https://github.com/jiangwen365/pypyodbc) para la conectividada con las bases de datos: `pip install pypyodbc`
		* [pyinstaller](https://github.com/pyinstaller/pyinstaller/) solo si el objetivo final es construi un ejecutable binario, esta herramienta es bastante sencilla y rápida si bien es mucho más poderosa [Cx_freeze](https://bitbucket.org/anthony_tuininga/cx_freeze=: `pip install pyinstaller`


* Probar el xls2table
	* Activar el entorno:  `source  ../venvs/xls2table/Scripts/activate`
	* Ejecuta el script principal: `python xls2table.py -h`


* Preparar EXE para distribución
	* `pyinstaller xls2table.py -y --onefile`
	* El archivo final debería estar en ./dist/xls2table.exe



Documentación
=============

```
#!bash

uso: xls2table [-h] [-v] [-l <level>] [-n <numero>] [-c] [-t]
               inputfile outputtable dsn

Importa datos de Excel a una tabla SQL 2016, Patricio Moracho
<pmoracho@gmal.com>

argumentos posicionales:
  inputfile                         Archivo Excel de entrada
  outputtable                       Tabla SQL donde se insertarán las filas de
                                    la planilla
  dsn                               Cadena de conexión

argumentos opcionales:
  -h, --help                        mostrar esta ayuda y salir
  -v, --version                     Mostrar el número de versión y salir
  -l <level>, --log <level>         Nivel de log
  -n <numero>, --sheetnum <numero>  Número de solapa a leer (la primera es 0)
  -c, --hasheader                   La planilla tiene en la primer fila el
                                    nombre de los campos?
  -t, --nativetypes                 Intentar respetar el tipo de dato de cada
                                    columna, sino los campo por defecto se
                                    generan como VARCHAR(255)
```

Construcción de la cadena Dsn según datasource
==============================================

	* SQL Server: "DRIVER={SQL Server};SERVER=<server>;DATABASE=<database>;UID=<usuario>;PWD=<password>" 

Niveles de log
==============

Utilizar el parámetro `-l` o `--log` para indicar el nivel de información que mostrará la herramienta. Por defecto el nivel es NONE, que no mustra ninguna información.

Nível		| Detalle
----------- | -------------
NONE		| No motrar ninguna información
DEBUG		| Información detallada, tipicamente análisis y debug
INFO		| Confirmación visual de lo esperado
WARNING		| Información de los eventos no esperados, pero aún la herramienta puede continuar
ERROR		| Errores, alguna funcionalidad no se puede completar
CRITICAL 	| Errores serios, el programa no puede continuar

