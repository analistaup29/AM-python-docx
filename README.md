# AM-python-docx
Este repositorio contiene el código para el proyecto de automatización de la AM para regiones y Unidades Ejecutoras utilizando python-docx.

## Contenido

-[Scripts](https://github.com/vladi3101/AM-python-docx/tree/main/scripts)
-[Input](https://github.com/vladi3101/AM-python-docx/tree/main/input)
-[Output](https://github.com/vladi3101/AM-python-docx/tree/main/output)

## Estructura del proyecto

Se cuentan con tres scripts principales `01_AM_region_corta.py` contiene el código para actualizar la AM a nivel regional. `02_AM_region_anexos.py` contiene el código para actualizar los anexos de la AM regional y `03_AM_unidad_ejecutora.py` contiene el código de la AM a nivel unidad ejecutora.

Las `01_AM_region_corta.py` y `02_AM_region_anexos.py` contienen información de las Intervenciones Pedagógicas, Compromisos de desempeño, Mascarillas y protectores faciales, Pago de Encargaturas, Pago de asignaciones temporales, Pago de beneficios sociales, Financiamiento de plazas 2021 y Deudad sociales. El script `03_AM_unidad_ejecutora.py` solo contiene información de las Intervenciones Pedagógicas.

## Replicar y/o actualizar el repositorio

Aquí se incluyen los pasos para replicar o actualizar los scripts que generan las AM automatizadas. Si encuentras algún problema para correr el código o reproducir los resultados, por favor [crea un `Informe de problemas`](https://github.com/analistaup29/00_Data/issues/new) en este repositorio.

### Requerimientos de software

- Python (código se corrió con la versión 3.8)
  -  import docx
  -  import pandas as pd
  -  import numpy as np
  -  import nums_from_string
  -  import os
  -  import getpass
  -  import glob
  -  import matplotlib.pyplot as plt
  -  from datetime import datetime
  -  import pyodbc
  -  from janitor import clean_names 
  -  from pathlib import Path
  -  from docx.shared import Pt
  -  from docx.shared import Inches

### Instrucciones para replicar

1. Clona el repositorio en tu PC Minedu
2. Ve a la carpeta `scripts`, allí encontrarás los scripts llamados `01_AM_region_corta.py`, `02_AM_region_anexos.py` y `03_AM_unidad_ejecutora.py`. Estos scripts te permiten actualizar las AMs. 
3. Todas las bases se guardan en la Unidad B de Minedu y son accesibles desde el sharepoint o desde el VPN virtual. Ve a `unidad_B/2021/4. Herramientas de Seguimiento/13.AM_automatizada/Input`, allí encontrarás los inputs o archivos que utiliza la AM.
4. Actualiza las bases y corre cualquiera de los scripts que deba ser actualizado. No olvides cambiar la ruta de usuario dentro de cada script.
