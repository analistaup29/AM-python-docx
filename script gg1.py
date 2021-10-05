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
    
    ##Párrafos
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
    encarg_parrafo2.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
    
    encarg_parrafo3 = document.add_paragraph('Con ')
    encarg_parrafo3.add_run('Decreto Supremo 217-2021 \
publicado el 27 de agosto de 2021 en el marco de lo autorizado en el literal b) \
del numeral 40.1 de la Ley de Presupuesto 2021, se ha realizado una transferencia \
de partidas por el monto de S/.')  # Este párrafo tendrá que variar año tras año
    encarg_parrafo3.add_run(' XXXXXXX') #Insertar valor de base de datos    
    encarg_parrafo3.add_run(' a favor de las Unidades Ejecutoras de Educación de la Región ')
    encarg_parrafo3.add_run(region)
    encarg_parrafo3.add_run(' para financiar el costo diferencial.')
    encarg_parrafo3.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    encarg_parrafo4 = document.add_paragraph('En el siguiente cuadro se muestran el costo \
y los montos programados/transferidos a la Región ')
    encarg_parrafo4.add_run(region)
    encarg_parrafo4.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

#----------------------------------------------------------------------------------------------#
    document.add_heading("2.Pago de Asignaciones Temporales", level=1)
    
    ##Párrafos
    encarg_parrafo5 = document.add_paragraph('Para la Región ')    
    encarg_parrafo5.add_run(region)
    encarg_parrafo5 = document.add_paragraph(', por concepto de Asignaciones Temporales \
por prestar servicios en condiciones especiales, se ha calculado para el 2021 un costo de \
S/ ')    
    encarg_parrafo5.add_run(' XXXXXXX') #Insertar valor de base de datos    
    encarg_parrafo5 = document.add_paragraph('que incluye el pago por prestar servicios \
en zonas rurales, de frontera, VRAEM, Instituciones Educativas Unidocentes, Multigrado \
Bilingüe y acreditar dominio de lengua originaria, de los profesores y auxiliares de \
educación nombrados y contratados.')        
    encarg_parrafo5.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    
    encarg_parrafo6 = document.add_paragraph('Para financiar estos conceptos, el ')
    encarg_parrafo6.add_run('2020 ') #Calcular valor de año anterior
    encarg_parrafo6.add_run('el MINEDU gestionó una programación directa de recursos \
en el PIA ')
    encarg_parrafo6.add_run(datetime.today().strftime('%Y')) #Año actual
    encarg_parrafo6.add_run(' de las Unidades Ejecutoras de Educación de la Región ')    
    encarg_parrafo6.add_run(region)
    encarg_parrafo6.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
  
    encarg_parrafo7 = document.add_paragraph('Con Decreto Supremo 187-2021 \
publicado el 22 de julio de 2021 en el marco de lo autorizado en los literales \
a), c), d) y e) del numeral 40.1 de la Ley de Presupuesto 2021, ') # Este párrafo tendrá que variar año tras año
    encarg_parrafo7.add_run('se ha realizado una transferencia \
de partidas por el monto de S/. ')
    encarg_parrafo7.add_run(' XXXXXXX') #Insertar valor de base de datos    
    encarg_parrafo7.add_run(' a favor de las Unidades Ejecutoras de Educación \
de la Región ') 
    encarg_parrafo7.add_run(region)
    encarg_parrafo7.add_run(' para financiar el costo diferencial de las asignaciones \
temporales a favor los profesores y auxiliares de educación nombrados y contratados.')    
    encarg_parrafo7.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    encarg_parrafo8 = document.add_paragraph('Actualmente está en gestión en el MINEDU \
la segunda transferencia de recursos por concepto de asignaciones temporales, \
el cual debería realizarse antes del ')
    encarg_parrafo8.add_run('26 de noviembre del 2021') #Esta fecha se actualizará año a año 
    encarg_parrafo8.add_run(' de acuerdo al plazo legal establecido en la Ley de Presupuesto') 
    encarg_parrafo8.add_run(datetime.today().strftime('%Y')) #Año actual
    encarg_parrafo8.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
    
    encarg_parrafo9 = document.add_paragraph('En el siguiente cuadro se muestran el costo y los \
montos programados/transferidos a la Región ')
    encarg_parrafo9.add_run(region)
    encarg_parrafo9.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

#----------------------------------------------------------------------------------------------#

    document.add_heading("3.Pago de Beneficios Sociales", level=1)

    ##Párrafos    
    encarg_parrafo10 = document.add_paragraph('Para la Región ')
    encarg_parrafo10.add_run(region)
    encarg_parrafo10.add_run(', de acuerdo con la nueva estrategia para el pago oportuno \
de Beneficios Sociales implementada por el MINEDU, se han aprobado pagos por concepto de \
Asignación por Tiempo de Servicios (ATS), Compensación por Tiempo de Servicios (CTS) y, \
Subsidio por Luto y Sepelio (SLS) hasta por  un costo de S/ ')
    encarg_parrafo10.add_run(' XXXXXXX') #Insertar valor de base de datos    
    encarg_parrafo10.add_run(' a la fecha , a favor de los profesores y auxiliares de educación \
nombrados y contratados.')    
    encarg_parrafo10.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    encarg_parrafo11 = document.add_paragraph('Para financiar estos conceptos, el ')
    encarg_parrafo11.add_run('2020 ') #Calcular valor de año anterior
    encarg_parrafo11.add_run('el MINEDU gestionó una programación directa de recursos en el PIA ')
    encarg_parrafo11.add_run(datetime.today().strftime('%Y')) #Año actual
    encarg_parrafo11.add_run(' de las Unidades Ejecutoras de Educación de la Región ')
    encarg_parrafo11.add_run(region)
    encarg_parrafo11.add_run(' por el monto de S/. ')
    encarg_parrafo11.add_run(' XXXXXXX.') #Insertar valor de base de datos    
    encarg_parrafo11.add_run('  Lo cual fue comunicado a través del ')
    encarg_parrafo11.add_run('Oficio Múltiple N° 00011-2021-MINEDU/SPE-OPEP-UPP') #Esto cambiará cada año
    encarg_parrafo11.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    encarg_parrafo12 = document.add_paragraph('Con ')
    encarg_parrafo12.add_run('Decreto Supremo 072-2021 publicado el 21 de abril de 2021')
    encarg_parrafo12.add_run('en el marco de lo autorizado en los literales a), d) y e) \
del numeral 40.1 de la Ley de Presupuesto ')
    encarg_parrafo12.add_run(datetime.today().strftime('%Y')) #Año actual
    encarg_parrafo12.add_run(', se ha realizado una transferencia de partidas por el monto de S/ ')   
    encarg_parrafo12.add_run(' XXXXXXX.') #Insertar valor de base de datos   
    encarg_parrafo12.add_run(', a favor de las Unidades Ejecutoras de Educación de la Región ')   
    encarg_parrafo12.add_run(region)
    encarg_parrafo12.add_run(' para financiar el pago de los beneficios sociales a favor los profesores \
y auxiliares de educación nombrados y contratados que fueron reconocidos hasta el ')   
    encarg_parrafo12.add_run('2020') #Calcular valor de año anterior
    encarg_parrafo12.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    encarg_parrafo13 = document.add_paragraph('Asimismo, mediante ')
    encarg_parrafo13.add_run('Decreto Supremo 256-2021 publicado el 24 de setiembre de 2021, ')  #Esto cambiará cada año
    encarg_parrafo13.add_run(', se realizó la segunda transferencia de recursos por concepto de \
beneficios sociales a favor de docentes y auxiliares nombrados y contratados, cuyos beneficios fueron \
reconocidos durante el año ')
    encarg_parrafo13.add_run(datetime.today().strftime('%Y')) #Año actual
    encarg_parrafo13.add_run(' transfiriéndose S/. ')
    encarg_parrafo13.add_run(' XXXXXXX.') #Insertar valor de base de datos   
    encarg_parrafo13.add_run('  a las Unidades Ejecutoras de Educación de la Región ')
    encarg_parrafo13.add_run(region)
    encarg_parrafo13.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
   

    encarg_parrafo14 = document.add_paragraph('En el siguiente cuadro se muestran el costo y los montos \
programados/transferidos a la Región ')
    encarg_parrafo14.add_run(region)
    encarg_parrafo14.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    encarg_parrafo15 = document.add_paragraph(' De la misma forma, durante el presente año, para la Región ')  
    encarg_parrafo15.add_run(region)
    encarg_parrafo15.add_run(' se ha realizado las siguientes transferenciasn')
    encarg_parrafo15.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    encarg_parrafo16 = document.add_paragraph(' Por otro lado, para el año ')
    encarg_parrafo16.add_run('2022 ') #Calcular año posterior
    encarg_parrafo16.add_run('el MINEDU está gestionando la programación parcial de recursos en los \
presupuestos de las Unidades Ejecutoras para atender encargaturas, asignaciones temporales, \
beneficios sociales, entre otros y, el financiamiento restante, se realizará de manera oportuna el ')
    encarg_parrafo16.add_run('2022, ') #Calcular año posterior
    encarg_parrafo16.add_run('preferentemente antes que termine el primer semestre de dicho año fiscal.')    
    encarg_parrafo16.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

#------------------------------------------------------------------------------------------------------------#
#------------------------------------------------------------------------------------------------------------#    
    # Incluimos sección 2 de Proceso de racionalización
    document.add_heading("Sobre el proceso de racionalización", level=1)
    run.underline = True    

    document.add_heading("4. Financiamiento de plazas 2021, en el marco del \
proceso de racionalización 2020", level=1)

    #Párrafo 
    racio_parrafo1 = document.add_paragraph('A través del ',  style="List Bullet")
    racio_parrafo1.add_run('DS 078-2021-EF') #Esto cambia año tras año
    racio_parrafo1.add_run(' se financiaron ') 
    racio_parrafo1.add_run(' XXXXXXX') #Insertar valor de base de datos 
    racio_parrafo1.add_run(' plazas de docentes de aula en el marco de los resultados del \
proceso de racionalización ') 
    racio_parrafo1.add_run('2020 ') #Calcular valor de año anterior
    racio_parrafo1.add_run('en servicios educativos públicos de la región ')
    racio_parrafo1.add_run(region)
    racio_parrafo1.add_run(' con la siguiente distribución por UGEL:')
    racio_parrafo1.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY


    document.add_heading("5. Resultados proceso de racionalización 2020", level=1)
 
    #Párrafo 
    racio_parrafo2 = document.add_paragraph('En el proceso de racionalización ',  style="List Bullet")    
    racio_parrafo2.add_run('2020 ') #Calcular valor de año anterior
    racio_parrafo2.add_run(', se identificó en la región ')
    racio_parrafo2.add_run(region)
    racio_parrafo2.add_run(' un total de ')
    negrita1 = racio_parrafo2.add_run('XXXXXXX') #Insertar valor de base de datos
    negrita1.bold= True
    negrita2 = racio_parrafo2.add_run(' plazas de docentes de aula excedentes ')
    negrita2.bold= True
    racio_parrafo2.add_run('y ')
    negrita3 = racio_parrafo2.add_run('XXXXXXX') #Insertar valor de base de datos
    negrita3.bold= True    
    negrita4 = racio_parrafo2.add_run(' plazas de requerimiento. ')
    negrita4.bold= True
    racio_parrafo2.add_run('A partir de esos resultados, se procedió a calcular \
el requerimiento y la excedencia por UGEL y el agregado a nivel regional, \
ello se puede observar en las dos últimas columnas del siguiente cuadro: ')  
    racio_parrafo2.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    #(Insertar cuadro)

    racio_parrafo3 = document.add_paragraph(' Por lo tanto, a nivel regional \
se contaba con una brecha interna de ',  style="List Bullet")    
    racio_parrafo3.add_run('XXXXXXX') #Insertar valor de base de datos
    racio_parrafo3.add_run(' plazas en ')
    racio_parrafo3.add_run('XXXXXXX') #Insertar valor de base de datos
    racio_parrafo3.add_run(' UGEL, y un excedente neto de \
plazas vacantes ascendente')
    racio_parrafo3.add_run(' XXXXXXX') #Insertar valor de base de datos
    racio_parrafo3.add_run(' plazas en ')     
    racio_parrafo3.add_run(' XXXXXXX') #Insertar valor de base de datos
    racio_parrafo3.add_run(' UGEL. Con ello, se obtuvo un ')      
    negrita5 = racio_parrafo3.add_run('requerimiento neto a nivel regional igul a ')
    negrita5.bold= True
    negrita6 = racio_parrafo3.add_run('XXXXXXX') #Insertar valor de base de datos
    negrita6.bold= True     
    racio_parrafo3.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    document.add_heading("6. Acciones de reordenamiento territorial 2020", level=1)
 
    #Párrafo
    racio_parrafo4 = document.add_paragraph('En el marco del proceso de racionalización ',  style="List Bullet") 
    racio_parrafo4.add_run('2020 ') #Calcular valor de año anterior
    racio_parrafo4.add_run(', en la región ')
    racio_parrafo4.add_run(region)    
    racio_parrafo4.add_run(' no se inhabilitaron plazas a pesar de contar con ')
    racio_parrafo4.add_run(' XXXXXXX') #Insertar valor de base de datos
    racio_parrafo4.add_run(' plazas vacantes identificadas como excedencia neta.')
    racio_parrafo4.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

#------------------------------------------------------------------------------------------------------------#
#------------------------------------------------------------------------------------------------------------#    
    # Incluimos sección 3 deudas sociales y contratación
    document.add_heading("Sobre Deudas Sociales y Contratación del DL 276", level=1)
    run.underline = True  

    document.add_heading("7. Deudas Sociales", level=1)

    #Párrafo
    deuda_parrafo1 = document.add_paragraph(' La Deuda Social del Sector Educación \
está constituida por las obligaciones de pago  en materia laboral y previsional, \
respecto del personal Docente y Auxiliares de Educación provenientes \
de la derogada Ley del Profesorado  y su modificatoria; \
del personal Administrativo sujeto al Decreto Legislativo Nº 276, \
Ley de Bases de la Carrera Administrativa; del Personal Administrativo \
en el marco del Texto Único Ordenado (TUO) del Decreto Legislativo Nº 728 \
Ley de Productividad y Competitividad  Laboral  aprobado por \
Decreto Supremo Nº 003-97-TR. ') 
    deuda_parrafo1.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    deuda_parrafo2 = document.add_paragraph(' La atención de la deuda social \
con el Sector se viene efectuando a través de transferencias de \
partidas del Tesoro Público; es decir, en forma complementaria a \
los recursos presupuestales con que disponen los Pliegos del Gobierno Nacional \
y los Gobiernos Regionales, para la atención del pago de sentencias judiciales \
en calidad de cosa juzgada y en ejecución. Dicho tratamiento, \
se realiza desde el año 2014 en el marco de las Leyes Anuales de Presupuesto, \
habiéndose autorizado diversas habilitaciones presupuestarias \
en el nivel institucional mediante los decretos supremos correspondientes. ')
    deuda_parrafo2.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    deuda_parrafo3 = document.add_paragraph(' Desde el año 2018 se han venido \
destinando S/ 200 000 000,00 (DOSCIENTOS MILLONES y 00/100 DE SOLES) \
para el sector Educación, que según los criterios de priorización aprobados \
por el Ministerio de Educación, son destinados preferentemente \
al pago de las sentencias judiciales que reconocen \
el derecho de preparación de clases, frente a otros conceptos. ')
    deuda_parrafo3.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    deuda_parrafo4 = document.add_paragraph(' En ese sentido, mediante el numeral \
1.2 del artículo 1 del Decreto Supremo N° 216-2021-EF, publicado en el \
Diario Oficial “El Peruano” el 27 de agosto de 2021, se autorizó \
la Transferencia de Partidas en el Presupuesto del Sector Público \
para el Año Fiscal 2021, hasta por la suma de S/ 200 000 000,00 \
(DOSCIENTOS MILLONES Y 00/100 SOLES), a favor de diversos Pliegos \
del Gobierno Nacional (dentro de los que se encuentra el Ministerio de Educación \
(MINEDU), y los Gobiernos Regionales, para financiar el pago de \
sentencias judiciales en calidad de cosa juzgada del sector Educación \
y en ejecución al 31 de diciembre de 2020, en el marco del numeral 6 de \
la de la Undécima Disposición Complementaria Final de la Ley N° 31084, \
Ley de presupuesto del año fiscal 2021, con cargo a los recursos de la Reserva \
de Contingencia del Ministerio de Economía y Finanzas. \
El detalle de dicha transferencia de recursos se muestra a continuación:')
    deuda_parrafo4.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    #(Insertar cuadro)

    deuda_parrafo5 = document.add_paragraph('Del mismo modo, es importante \
señalar que las deudas sociales del Sector Educación se atienden siguiendo \
los criterios de priorización establecidos en el presente año fiscal mediante \
el Decreto Supremo N° 003-2021-MINEDU, Decreto Supremo que aprueba \
los criterios de priorización para la atención del pago de sentencias judiciales \
en calidad de cosa juzgada y en ejecución del Sector Educación, \
donde no se establece una distinción entre regímenes laborales \
o entre personal activo o cesante, sino se siguen patrones de antigüedad, \
salud, entre otros')
    deuda_parrafo5.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY


    deuda_parrafo6 = document.add_paragraph('Es igualmente importante señalar, \
que en el caso de existir déficit presupuestal durante el ejercicio fiscal \
en ejecución, para la atención de las planillas continuas de pensionistas, \
derivadas de sentencias en calidad de cosa juzgada, \
su financiamiento deberá tener en cuenta la exoneración prevista \
en el numeral 9.2 del artículo 9. Medidas en materia de modificaciones \
Presupuestarias en el Nivel Funcional Programático, de la Ley Nº 31084, \
Ley de Presupuesto del Sector Público para el Año Fiscal 2021; \
en tanto prescribe que: ')

    formato1 = deuda_parrafo6.add_run(' “A nivel de pliego')
    formato1.italic = True
    formato2 = deuda_parrafo6.add_run(' la Partida de Gasto 2.2.1 “Pensiones” no puede ser habilitadora, ')
    formato2.italic = True
    formato2.bold = True
    formato3 = deuda_parrafo6.add_run('salvo para las habilitaciones que se realicen ')
    formato3.italic = True
    formato3.bold = True
    formato3.underline = True
    formato4 = deuda_parrafo6.add_run('dentro de la misma partida entre unidades \
ejecutoras del mismo pliego presupuestario, y ')
    formato4.italic = True
    formato4.bold = True    
    formato5 = deuda_parrafo6.add_run('para la atención de sentencias judiciales \
en materia pensionaria con calidad de cosa juzgada,')
    formato5.italic = True
    formato5.bold = True
    formato5.underline = True    
    formato6 = deuda_parrafo6.add_run(' en este último caso, previo informe \
favorable de la Dirección General de Presupuesto Público (DGPP), ')
    formato6.italic = True
    formato6.bold = True  
    formato7 = deuda_parrafo6.add_run(' y de corresponder, sobre la base \
de la información registrada en el Aplicativo Informático para el \
 Registro Centralizado de Planillas y de Datos de los Recursos Humanos \
 del Sector Público (AIRHSP) que debe remitir la Dirección General \
 de Gestión Fiscal de los Recursos Humanos a la DGPP.(…)”.')
    formato7.italic = True
    deuda_parrafo6.add_run('(Lo resaltado es nuestro).')
    deuda_parrafo6.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    deuda_parrafo7 = document.add_paragraph('En el mismo numeral del artículo \
acotado en el párrafo anterior, se establece que las solicitudes de \
informe favorable para su aplicación solo pueden ser presentadas al \
Ministerio de Economía y Finanzas hasta el 15 de octubre de 2021; por lo que, ')
    formato8 = deuda_parrafo7.add_run('las unidades orgánicas pertinentes \
del Pliego Presupuestal involucrado, deben implementar \
las acciones administrativas a que hubiera lugar con sujeción \
a la normatividad presupuestaria y al plazo legal invocados líneas atrás.')
    formato8.bold = True
    deuda_parrafo7.add_run(' Así como, ')
    formato9 = deuda_parrafo7.add_run('de corresponder, \
las acciones pertinentes para el pago de las deudas que se hubieran generado \
por aplicación de la sentencia judicial en calidad de cosa juzgada \
que así lo ordene, en cuyo caso deberá observarse lo dispuesto en \
la Ley Nº 30137 y su Reglamento.')
    formato9.underline = True
    deuda_parrafo7.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    document.add_heading("8. Ruta para contratación de personal administrativo", level=1)

    #Párrafo
    pers_adm_parrafo1 = document.add_paragraph('En el marco del inciso \
d) del numeral 8.1 del artículo 8 de la Ley N° 31084, \
Ley de Presupuesto del Sector Público para el año 2021, ')
    formato10 = pers_adm_parrafo1.add_run('no se ha establecido como excepción \
la posibilidad del ingreso de personal administrativo \
bajo el régimen laboral del Decreto Legislativo N° 276.')
    formato10.underline = True
    pers_adm_parrafo1.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
    
    pers_adm_parrafo2 = document.add_paragraph('A su vez, \
la Gerencia de Políticas de Gestión del Servicio Civil de \
la Autoridad Nacional del Servicio Civil (SERVIR) se ha pronunciado \
sobre la imposibilidad de contratar personal administrativo del \
Decreto Legislativo N° 276, emitido mediante Informe Técnico \
N° 331-2021-SERVIR-GPGSC.')
    pers_adm_parrafo2.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    pers_adm_parrafo3 = document.add_paragraph('Dado que el Decreto Legislativo \
N° 276, no solo rige al Sector Educación, sino a todo el Sector Público, \
desde el Ministerio de Educación se propuso incluir la habilitación \
para su contratación en una norma con rango de Ley, la cual autorizaría \
la contratación de personal del Decreto Legislativo N° 276. \
Sin embargo, dicha propuesta no fue incluida en el \
Proyecto de Ley del Presupuesto del Sector Público para el año 2022, \
enviada al Congreso de la República para su aprobación el 31 de agosto de 2021.')
    pers_adm_parrafo3.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

    document.save(here() / f'output/{region}_AM_GG1.docx')
    