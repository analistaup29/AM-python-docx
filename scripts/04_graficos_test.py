# -*- coding: utf-8 -*-
"""
Created on Thu Dec  9 14:29:10 2021

@author: analistaup29
"""

import docx
from docx import Document
from docx.shared import Cm
import pandas as pd
from datetime import datetime
import os

# Creación del documento en base al template
document = docx.Document("formato.docx")

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

document.save('AM.docx')