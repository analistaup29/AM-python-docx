#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Sep 21 15:54:48 2021
"""

# Importar librerías ----------------------------------------------------------

import docx
from docx import Document
from docx.shared import Cm
import pandas as pd
from datetime import datetime
import os
from pyprojroot import here

# Acciones --------------------------------------------------------------------

# Creación del documento en base al template
document = docx.Document(here() / "input/formato.docx")

#Reemplazar con un loop
region = "Cajamarca"

#Título
title=document.add_heading('AM Temas Presupuestales', 0)
run = title.add_run()
run.add_break()
title.add_run('Región ')
title.add_run(region)
run = title.add_run()
run.add_break()
title.add_run(datetime.today().strftime('%d-%m-%y'))

document.add_heading('1. Financiamiento de plazas 2021', level=1)
pa1=document.add_paragraph('A través del DS 078-2021-EF, se financiaron ', style='List Bullet')
pa1.add_run('XXX') #Reemplazar 
pa1.add_run(' plazas de docentes de aula en el marco de los resultados del proceso de racionalización 2020 \
en servicios educativos públicos de la región ') 
pa1.add_run(region) 
pa1.add_run(' con la siguiente distribución por UGEL:') 


document.add_heading('2. Resultados proceso de racionalización 2020', level=1)
pa2= document.add_paragraph('En el proceso de racionalización 2020, se identificó en la región ', style='List Bullet') 
pa2.add_run(region).bold = True
pa2.add_run(' un total de ').bold = True
pa2.add_run('XXX').bold = True #Reemplazar 
pa2.add_run(' plazas de docentes de aula excedentes y ').bold = True
pa2.add_run('XXX').bold = True #Reemplazar 
pa2.add_run(' plazas de requerimiento').bold = True
pa2.add_run('. A partir de esos resultados, se procedió \
a calcular el requerimiento y la excedencia por UGEL y el agregado \
a nivel regional, ello se puede observar en las dos últimas columnas \
del siguiente cuadro:')

pa3= document.add_paragraph('En el proceso de racionalización 2020, se identificó en la región ', style='List Bullet') 
pa3.add_run(region).bold = True
pa3.add_run(' un total de ')
pa3.add_run('XXX') #Reemplazar 
pa3.add_run(' plazas de docentes de aula excedentes y ')
pa3.add_run('XXX') #Reemplazar 
pa3.add_run(' plazas de requerimiento. A partir de esos resultados, se procedió a  \
calcular el requerimiento y la excedencia por UGEL y el agregado a nivel regional,  \
ello se puede observar en las dos últimas columnas del siguiente cuadro:') 

pa4= document.add_paragraph('Por lo tanto, a nivel regional se contaba con una brecha interna de ', style='List Bullet')
pa4.add_run('XXX') #Reemplazar 
pa4.add_run(' plazas en ')
pa4.add_run('XXX') #Reemplazar 
pa4.add_run(' UGEL, y un excedente neto de plazas vacantes ascendente a ')
pa4.add_run('XXX') #Reemplazar 
pa4.add_run(' plazas en ')
pa4.add_run('XXX') #Reemplazar 
pa4.add_run(' UGEL. Con ello, se obtuvo un requerimiento neto a nivel regional igual a ')
pa4.add_run('XXX') #Reemplazar 
pa4.add_run('.') 


document.add_heading('3. Acciones de reordenamiento territorial 2020', level=1)

pa5= document.add_paragraph('En el marco del proceso de racionalización 2020, en la región ', style='List Bullet')
pa5.add_run(region)
pa5.add_run(' no se inhabilitaron plazas a pesar de contar con ')
pa5.add_run('XX')
pa5.add_run(' plazas vacantes identificadas como excedencia neta.') 


document.add_heading('4. Intervenciones pedagógicas ', level=1)
pa6= document.add_paragraph('En las Unidades Ejecutoras de Educación de ', style='List Bullet')
pa6.add_run(region)
pa6.add_run(' vienen implementando ')
pa6.add_run('XXX') #Reemplazar 
pa6.add_run(' intervenciones y acciones pedagógicas en el Año 2021, en el marco de la Norma  \
Técnica “Disposiciones para la implementación de las intervenciones y acciones pedagógicas  \
del Ministerio de Educación en los Gobiernos Regionales y Lima Metropolitana en el Año Fiscal  \
2021”, aprobada mediante RM N° 043-2021-MINEDU y modificada RM N° 159-2021-MINEDU.') 
                       
pa7= document.add_paragraph('Las Unidades Ejecutoras de Educación de ', style='List Bullet')
pa7.add_run(region)
pa7.add_run(' cuentan con S/ ')
pa7.add_run('XXX') #Reemplazar 
pa7.add_run(' en su Presupuesto Institucional Modificado (PIM), que permite cubrir el costo \
para la implementación de las referidas intervenciones y acciones pedagógicas, por un monto de S/ ')
pa7.add_run('XXX') #Reemplazar 
pa7.add_run('. De dichos recursos, a fecha de corte del ')
pa7.add_run(datetime.today().strftime('%d-%m-%y'))
pa7.add_run(', se han ejecutado S/ ')
pa7.add_run('XXX') #Reemplazar 
pa7.add_run(', lo que equivale al ')
pa7.add_run('XXX') #Reemplazar 
pa7.add_run(' respecto al costo de los recursos.') 


# Tablas1
titu_tabla1= document.add_heading('Ejecución de intervenciones en ', level=1)
titu_tabla1.add_run(region)

data = pd.read_excel(here() / "input/gaaa.xlsx")

t = document.add_table(data.shape[0]+1, data.shape[1])

for j in range(data.shape[-1]):
    t.cell(0,j).text = data.columns[j]

# add the rest of the data frame
for i in range(data.shape[0]):
    for j in range(data.shape[-1]):
        t.cell(i+1,j).text = str(data.values[i,j])
        
document.add_heading('5. Condiciones para la reapertura de IIEE: Mascarillas y protectores faciales', level=1)


document.save(here() / "output/AM.docx")        
        