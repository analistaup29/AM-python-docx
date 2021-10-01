# -*- coding: utf-8 -*-
"""
Created on Thu Sep 30 10:51:08 2021
"""

# Importar librerías ----------------------------------------------------------
import docx
import pandas as pd
from datetime import datetime
from pyprojroot import here # pip install pypro
from janitor import clean_names # pip install pyjanitor

# Opciones --------------------------------------------------------------------

# Formato de tablas
pd.options.display.float_format = '${:,.0f}'.format

# Transformación de Datasets --------------------------------------------------

# Sobre el financiamiento de conceptos remunerativos
# C) Base de Encargaturas

# D) Base de Asignaciones Temporales

# E) Base de Beneficios Sociales

# Sobre el proceso de racionalización




# Generamos la lista de Regiones
lista_regiones = ["AMAZONAS"]

# For loop para cada región
for region in lista_regiones:
    
    
    #########################################################################
    # Incluimos el código del Documento
    document = docx.Document(here() / "input/Formato.docx") # Creación del documento en base al template
    title=document.add_heading('AM TEMAS PRESUPUESTALES - REGION ', 0) #Título del documento
    title.add_run(region)
    run = title.add_run()
    run.add_break()
    run = title.add_run()
    run.add_break()
    title.add_run(datetime.today().strftime('%d-%m-%y'))
    # Incluimos sección 1 de intervenciones pedagógicas
    document.add_heading("Sobre el financiamiento de conceptos remunerativos", level=1)
    run.underline = True
    document.add_heading("1.Pago de Encargaturas", level=1)
    encarg_parrafo1 = document.add_paragraph(
    "Para la región ")
    encarg_parrafo1.add_run(region)
    encarg_parrafo1.add_run(', por concepto de encargaturas, se ha calculado para el ')
    encarg_parrafo1.add_run(datetime.today().strftime('%Y'))
    encarg_parrafo1.add_run(', un costo de S/.')
    encarg_parrafo1.add_run(', XXXXXXX') #Insertar valor de base de datos
    encarg_parrafo1.add_run(
    ' que incluye la Jornada de Trabajo Adicional de 10 horas \
la carga social vinculada y la asignación por cargo de \
los profesores que asumen cargos de mayor responsabilidad \
mediante encargaturas')
    encarg_parrafo1.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
    
    encarg_parrafo2 = document.add_paragraph('Para financiar estos conceptos, el')
    #Insertar valor del año pasado
    encarg_parrafo2.add_run('el MINEDU gestionó una programación directa de recursos\
    en el PIA 2021 de las Unidades Ejecutoras de Educación de la Región ')
    encarg_parrafo2.add_run(region)  
    encarg_parrafo2.add_run(' por el monto de S/.')
    encarg_parrafo2.add_run(', XXXXXXX') #Insertar valor de base de datos
    encarg_parrafo2.add_run('  en la finalidad 0267929 Pago de la asignación por jornada \
de trabajo adicional y asignación por cargo de mayor responsabilidad, \
la cuál es usada para financiar las encargaturas. Asimismo, el Pliego Regional \
ya contaba con una programación de ')
    encarg_parrafo2.add_run(', XXXXXXX') #Insertar valor de base de datos
    encarg_parrafo2.add_run(' en la misma finalidad y mediante ')
    encarg_parrafo2.add_run(' Oficio Múltiple N° 00082-2021-MINEDU/SPE-OPEP-UPP, ')
    encarg_parrafo2.add_run('se le solicitó a las Unidades Ejecutoras del Pliego Regional \
#realizar modificaciones presupuestarias por el monto de S/.')
    encarg_parrafo2.add_run(' XXXXXXX') #Insertar valor de base de datos
    encarg_parrafo2.add_run(' para habilitar la finalidad 0267929')
    
    document.save(here() / f'output/{region}_AM_GG1.docx')
    