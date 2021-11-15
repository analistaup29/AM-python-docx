# -*- coding: utf-8 -*-
"""

"""

# Paquetes
import pyodbc
import pandas as pd

##############################################################################
# Cargamos data
##############################################################################

# Conectarse a BD UPP
cnxn = pyodbc.connect(driver='{SQL Server}', server='10.200.2.45', database='db_territorial_upp',
                      trusted_connection='yes')
cursor = cnxn.cursor()

# Cargamos data Disponibilidad
query = "SELECT * FROM dbo.disponibilidad_presupuestal;"
base_disponibilidad = pd.read_sql(query, cnxn)


