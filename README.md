# xls2table

Herramienta para convertir una planilla Excel (xls, xlsx) en un tabla SQL. La
lectura del formato es nativa y no requiere de Microsoft Excel instalado.

# Empecemos

Antes que nada, necesitaremos:

* [Git for Windows](https://git-scm.com/download/win) instalado y funcionando
* Una terminal de Windows, puede ser `cmd.exe`

Con **Git** instalado, desde la línea de comando y con una carpeta dónde
alojaremos este proyecto, por ejemplo `c:\proyectos`, simplemente:

``` 
c:\> c: 
c:\> cd \proyectos 
c:\> git clone <url del repositorio>
c:\> cd <carpeta del repositorio>
``` 

|                       |                                           |
| --------------------- |-------------------------------------------|
| Repositorio           | https://github.com/pmoracho/xls2table.git |
| Carpeta del proyecto  | .                                         |


## Instalación de **Python**

Para desarrollo de la herramienta es necesario, en primer término, descargar un
interprete Python. **xls2table** ha sido desarrollado con la versión 3.6, no es
mala idea usar esta versión, sin embargo debiera funcionar perfectamente bien
con cualquier versión de la rama 3x.

**Importante:** Si bien solo detallamos el procedimiento para entornos
**Windows**, el proyecto es totalmente compatible con **Linux**

* [Python 3.6.6 (32 bits)](https://www.python.org/ftp/python/3.6.6/python-3.6.6.exe)
* [Python 3.6.6 (64 bits)](https://www.python.org/ftp/python/3.6.6/python-3.6.6-amd64.exe)

Se descarga y se instala en el sistema el interprete **Python** deseado. A
partir de ahora trabajaremos en una terminal de Windows (`cmd.exe`). Para
verificar la correcta instalación, en particular que el interprete este en el `PATH`
del sistemas, simplemente corremos `python --version`, la salida deberá
coincidir con la versión instalada 

Es conveniente pero no mandatorio hacer upgrade de la herramienta pip: `python
-m pip install --upgrade pip`

## Instalación de `Virtualenv`

[Virutalenv](https://virtualenv.pypa.io/en/stable/). Es la herramienta estándar
para crear entornos "aislados" de **Python**. En nuestro ejemplo **xls2table**,
requiere de Python 3x y de varios "paquetes" adicionales de versiones
específicas. Para no tener conflictos de desarrollo lo que haremos mediante
esta herramienta es crear un "entorno virtual" en una carpeta del proyecto (que
llamaremos `venv`), dónde una vez "activado" dicho entorno podremos instalarle
los paquetes que requiere el proyecto. Este "entorno virtual" contendrá una
copia completa de **Python** y los paquetes mencionados, al activarlo se
modifica el `PATH` al `python.exe` que ahora apuntará a nuestra carpeta del
entorno y nuestras propias librerías, evitando cualquier tipo de conflicto con un
entorno **Python** ya existente. La instalación de `virtualenv` se hará
mediante:

```
c:\..\> pip install virtualenv
```

## Creación y activación del entorno virtual

La creación de nuestro entorno virtual se realizará mediante el comando:

```
C:\..\>  virtualenv venv --clear --prompt=[xls2table] --no-wheel
```

Para "activar" el entorno simplemente hay que correr el script de activación
que se encontrará en la carpeta `.\venv\Scripts` (en linux sería `./venv/bin`)

```
C:\..\>  .\venv\Scripts\activate.bat
[xls2table] C:\..\> 
```

Como se puede notar se ha cambiado el `prompt` con la indicación del entorno
virtual activo, esto es importante para no confundir entornos si trabajamos con
múltiples proyecto **Python** al mismo tiempo.

## Instalación de requerimientos

Mencionábamos que este proyecto requiere varios paquetes adicionales, la lista
completa está definida en el archivo `requirements.txt` para instalarlos en
nuestro entorno virtual, simplemente:

```
[xls2table] C:\..\> pip install -r requirements.txt
```

## Desarrollo

Si todos los pasos anteriores fueron exitosos, podríamos verificar si la
aplicación funciona correctamente mediante:

```
[xls2table] C:\..\> python xls2table.py
uso: xls2table [-h] [-v] [-l <level>] [-n <numero>] [-c] [-t] [-s]
               inputfile outputtable dsn
xls2table: error: los siguientes argumentos son requeridos: inputfile, outputtable, dsn

```

La ejecución sin parámetros arrojará la ayuda de la aplicación. A partir de
aquí podríamos empezar con la etapa de desarrollo.

## Generación del paquete para deploy

Para distribuir la aplicación en entornos **Windows** nos apoyaremos en
**Pyinstaller**, un modulo, instalado junto a los requerimientos, que nos
permite crear una carpeta de distribución de la aplicación totalmente portable.
Simplemente deberemos ejecutar el archivo `windist.bat`, al finalizar el
procesos deberías contar con una carpeta en `.\dist\xls2table` la cual será una
instalación totalmente portable de la herramienta, no haría falta nada más que
copiar la misma al equipo o servidor desde dónde deseamos ejecutarla.





* Preparar EXE para distribución
	* `pyinstaller xls2table.py -y --onefile`
	* El archivo final debería estar en ./dist/xls2table.exe



# Documentación

```
uso: xls2table [-h] [-v] [-l <level>] [-n <numero>] [-c] [-t] [-s]
               inputfile outputtable dsn

Importa datos de Excel a una tabla SQL (c) 2016, Patricio Moracho
<pmoracho@gmail.com>

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
  -s, --showonly                    Solo muestra el Script a ejecutar
```

## Construcción de la cadena Dsn según datasource

	* SQL Server: "DRIVER={SQL Server};SERVER=<server>;DATABASE=<database>;UID=<usuario>;PWD=<password>" 

## Niveles de log

Utilizar el parámetro `-l` o `--log` para indicar el nivel de información que
mostrará la herramienta. Por defecto el nivel es NONE, que no mustra ninguna
información.

Nível		| Detalle
----------- | -------------
NONE		| No motrar ninguna información
DEBUG		| Información detallada, tipicamente análisis y debug
INFO		| Confirmación visual de lo esperado
WARNING		| Información de los eventos no esperados, pero aún la herramienta puede continuar
ERROR		| Errores, alguna funcionalidad no se puede completar
CRITICAL 	| Errores serios, el programa no puede continuar

