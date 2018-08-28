@echo off

REM --------------------------------------------------------
REM Bat para la generación del paquete de deploy de xls2table
REM --------------------------------------------------------


REM --------------------------------------------------------
REM Creación del paquete para distribuir
REM --------------------------------------------------------
@echo --------------------------------------------------------
@echo Generando distribucion con pyinstaller..
@echo --------------------------------------------------------
@pyinstaller xls2table.py --onefile --noupx --clean --noconfirm

@echo --------------------------------------------------------
@echo Copiando archivos y herramientas adicionales..
@echo --------------------------------------------------------

REM --------------------------------------------------------
REM Eliminar archivos de trabajo
REM --------------------------------------------------------
@echo --------------------------------------------------------
@echo Eliminando archivos de trabajo ..
@echo --------------------------------------------------------
@rmdir build /S /Q
@del *.spec /S /F /Q

@echo --------------------------------------------------------
@echo Carpeta a distribuir dist\pboletin..
@echo --------------------------------------------------------
