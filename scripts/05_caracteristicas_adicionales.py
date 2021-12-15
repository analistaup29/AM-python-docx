#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Dec 15 14:11:52 2021

@author: bran

# Características adicionales para agregar a los scripts en python-docx

"""

###############################################################################
#1) Identificar automaticamente el documento 
#con fecha más reciente en una carpeta y obtener su fecha.
###############################################################################

## A) Base disponibilidad
# Importamos los nombres de los archivos en la carpeta intervenciones pedagogicas
#lista_archivos_int = glob.glob(os.path.join(proyecto,"input/intervenciones_pedagogicas/*"))
#Mantenemos el corte de disponibilidad más reciente
#fecha_corte_disponibilidad = max(lista_archivos_int, key=os.path.getctime)
# Nos quedamos con el nombre de archivo para la base de disponibilidad
#fecha_corte_disponibilidad = os.path.split(fecha_corte_disponibilidad)
#fecha_corte_disponibilidad = fecha_corte_disponibilidad[1]
# Extraemos la fecha del nombre de archivo
#fecha_corte_disponibilidad = nums_from_string.get_numeric_string_tokens(fecha_corte_disponibilidad)
# Convertimos a formato string
#fecha_corte_disponibilidad = ''.join(fecha_corte_disponibilidad) 
# Convertimos a formato numérico
#fecha_corte_disponibilidad_date = datetime.strptime(fecha_corte_disponibilidad, '%Y%m%d').date()
#mes_disponibilidad = fecha_corte_disponibilidad_date.month
# Damos estilo
#fecha_corte_disponibilidad_date = fecha_corte_disponibilidad_date.strftime("%d %b %Y")

############################################################################### 
#2) Conexión a SQL, permite conectarse directamente al SQL y descargar la data 
#desde allí.
###############################################################################

#cnxn = pyodbc.connect(driver='{SQL Server}', server='10.200.2.45', database='db_territorial_upp',
#                      trusted_connection='yes')
#cursor = cnxn.cursor()

# Cargamos data Disponibilidad
#query = "SELECT * FROM dbo.disponibilidad_presupuestal;"
#base_disponibilidad = pd.read_sql(query, cnxn)








