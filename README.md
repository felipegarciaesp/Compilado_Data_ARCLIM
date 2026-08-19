# Generador de compilatorio ARClim

Autor: Pablo Vergara

Descripción
-----------
Script para generar un archivo compilatorio en formato XLSX a partir de los archivos mensuales descargados desde ARClim. El script lee las tablas HTML dentro de archivos .xls mensuales, valida y concatena la información en una hoja final llamada "Compilado".

Archivo principal
- [`generar_compilatorio_tasmax_ssp585.py`](c:\Codigos\Compilado_Data_ARCLIM\generar_compilatorio_tasmax_ssp585.py)
  - Función de entrada: [`generar_compilatorio_tasmax_ssp585.main`](c:\Codigos\Compilado_Data_ARCLIM\generar_compilatorio_tasmax_ssp585.py)
  - Lectura de archivos mensuales: [`generar_compilatorio_tasmax_ssp585.read_month_file`](c:\Codigos\Compilado_Data_ARCLIM\generar_compilatorio_tasmax_ssp585.py)
  - Búsqueda de archivos por mes: [`generar_compilatorio_tasmax_ssp585.find_month_file`](c:\Codigos\Compilado_Data_ARCLIM\generar_compilatorio_tasmax_ssp585.py)
  - Construcción de la hoja compilada: [`generar_compilatorio_tasmax_ssp585.build_compiled_rows`](c:\Codigos\Compilado_Data_ARCLIM\generar_compilatorio_tasmax_ssp585.py)
  - Escritura del XLSX resultante: [`generar_compilatorio_tasmax_ssp585.write_xlsx`](c:\Codigos\Compilado_Data_ARCLIM\generar_compilatorio_tasmax_ssp585.py)
  - Variables relevantes: [`generar_compilatorio_tasmax_ssp585.VARIABLE_ESCENARIO`](c:\Codigos\Compilado_Data_ARCLIM\generar_compilatorio_tasmax_ssp585.py), [`generar_compilatorio_tasmax_ssp585.OUTPUT_DEFAULT`](c:\Codigos\Compilado_Data_ARCLIM\generar_compilatorio_tasmax_ssp585.py)

Estructura de carpeta esperada
------------------------------
Se asume una estructura como la siguiente:

- proyecto/
  - generar_compilatorio_tasmax_ssp585.py
  - mensual/                ← carpeta con los .xls mensuales (puede ser `.`, por defecto)
    - 01_*_ene_tasmax_ssp585.xls
    - 02_*_feb_tasmax_ssp585.xls
    - ...
    - 12_*_dic_tasmax_ssp585.xls

Por defecto el script toma como carpeta de entrada la carpeta actual. Puede especificarse con --input-dir.

Archivos mensuales
------------------
Los archivos Excel (en realidad tablas HTML dentro de .xls) que se muestran en ejemplos o en la carpeta son referenciales. Al ejecutar el script debe reemplazarlos por los archivos propios descargados desde ARClim que correspondan a cada mes.

Patrón de nombres
-----------------
El script busca archivos con el patrón:
`{mes:02d}_*_{mes_corto}_{VARIABLE_ESCENARIO}.xls`  
La variable usada para el escenario está en [`generar_compilatorio_tasmax_ssp585.VARIABLE_ESCENARIO`](c:\Codigos\Compilado_Data_ARCLIM\generar_compilatorio_tasmax_ssp585.py).

Archivo generado
----------------
Por defecto se genera:
- [`generar_compilatorio_tasmax_ssp585.OUTPUT_DEFAULT`](c:\Codigos\Compilado_Data_ARCLIM\generar_compilatorio_tasmax_ssp585.py) (00_IglesiaColorada_tasmax_ssp585_compilado.xlsx)  
La escritura del XLSX la realiza [`generar_compilatorio_tasmax_ssp585.write_xlsx`](c:\Codigos\Compilado_Data_ARCLIM\generar_compilatorio_tasmax_ssp585.py).

Uso
---
Ejemplo de ejecución:

```sh
python generar_compilatorio_tasmax_ssp585.py --input-dir mensual --output salida.xlsx
```

Notas finales
-------------
Queda pendiente averiguar si los archivos de cada mes deben entregarse con exactamente el formato de nombre que se observa en las planillas actuales; por ahora el script depende del patrón de nombres descrito arriba y dará error si no encuentra coincidencias.
