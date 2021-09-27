#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Sep 21 15:54:48 2021
"""

# Importar librerías ----------------------------------------------------------
import docx
import pandas as pd
from datetime import datetime
from pyprojroot import here
from janitor import clean_names # pip install pyjanitor

# Opciones --------------------------------------------------------------------

# Formato de tablas
pd.options.display.float_format = '${:,.0f}'.format

# Transformación de Datasets --------------------------------------------------

# A) Base de disponibilidad
## Cargamos base de disponibilidad
data_intervenciones = pd.read_excel(here() / "input/Disponibilidad_Presupuestal_20210923interv.xlsx")
data_intervenciones = clean_names(data_intervenciones) # Normalizamos nombres

# Mantenemos variables de interés (PIM, DEVENGADO, COMPROMETIDO CERTIFICADO) y 
# colapsamos a nivel de Region, Intervencion Pedagogica y Cas-No-Cas
tabla_intervenciones = data_intervenciones[["region", "cas_no_cas","intervencion_pedagogica", "pim_reporte_siaf_20210923", "presupuesto_certificado_reporte_siaf_20210923", "comprometido_anual_reporte_siaf_20210923", "presupuesto_devengado_reporte_siaf_20210923"]]. \
   groupby(by = ["region", "cas_no_cas", "intervencion_pedagogica"] , as_index=False).sum()

# Eliminamos filas de "No hay Intervenciones pedagogicas"
tabla_intervenciones = tabla_intervenciones[tabla_intervenciones['intervencion_pedagogica'] != "No hay Intervenciones Pedagógicas"]

# B) Siaf de mascarillas
## Cargamos la base insumo de mascarillas
data_mascarillas = pd.read_excel(here() / "input/Incorporación_DU_SIAF_20210921.xlsx", sheet_name='Sheet1')
data_mascarillas = clean_names(data_mascarillas) # Normalizamos nombres

# Mantenemos variables de interés (transferencia,  CERTIFICADO, COMPROMETIDO y DEVENGADO) y 
# colapsamos a nivel de Region y UE
tabla_mascarillas = data_mascarillas[["region","nom_ue","certificado","comprometido_anual","devengado","transferencia"]]. \
   groupby(by = ["region", "nom_ue"], as_index=False).sum()

tabla_mascarillas["region"] = tabla_mascarillas["region"].str.split(". ", n=1).apply(lambda l: "".join(l[1]))


# (PENDIENTE) FOR LOOP de número de intervenciones pedagógicas, Mascarillas y CDD
numero_intervenciones = "8" ## PENDIENTE
# Mascarillas y protectores faciales
fecha_corte_mascarillas = "21 de setiembre de 2021"
# Compromisos de desempeño
transferido_compromisos = "22,222"
acciones_centrales_cdd = "88,888"

# Generamos la lista de Regiones
lista_regiones = ["AMAZONAS", "TACNA", "AREQUIPA"]

# For loop para cada región
for region in lista_regiones:
    # Generamos la tabla "tabla1_region" - mantiene la región i de la lista de
    # regiones
    region_seleccionada = tabla_intervenciones['region'] == region
    tabla1_region = tabla_intervenciones[region_seleccionada]
    # Generamos los indicadores de PIM y ejecución de intervenciones
    pim_intervenciones_region = str('{:,.0f}'.format(tabla1_region["pim_reporte_siaf_20210923"].sum()))
    ejecucion_intervenciones_region = str('{:,.0f}'.format(tabla1_region["presupuesto_devengado_reporte_siaf_20210923"].sum()))
    # Generamos la tabla "tabla1_mascarilla" - mantiene la región i de la lista de
    # regiones
    region_seleccionada = tabla_mascarillas['region'] == region
    tabla2_region = tabla_mascarillas[region_seleccionada]
    # Generamos los indicadores de PIM y ejecución de intervenciones
    transferencia_mascarilla = str('{:,.1f}'.format(tabla2_region["transferencia"].sum()/1000000))
    devengado_mascarillas=str('{:.1%}'.format(tabla2_region["devengado"].sum()/tabla2_region["transferencia"].sum()))
    #########################################################################
    # Incluimos el código del Documento
    document = docx.Document(here() / "input/formato.docx") # Creación del documento en base al template
    title=document.add_heading('AYUDA MEMORIA', 0) #Título del documento
    run = title.add_run()
    run.add_break()
    title.add_run('REGIÓN ')
    title.add_run(region)
    run = title.add_run()
    run.add_break()
    title.add_run(datetime.today().strftime('%d-%m-%y'))
    # Incluimos sección 1 de intervenciones pedagógicas
    document.add_heading("1. Intervenciones pedagógicas", level=1) # 1) Intervenciones pedagógicas
    interv_parrafo1 = document.add_paragraph(
    "Las Unidades Ejecutoras de Educación de la región " , style="List Bullet")
    interv_parrafo1.add_run(region)
    interv_parrafo1.add_run(" vienen implementando ")
    interv_parrafo1.add_run(numero_intervenciones)
    interv_parrafo1.add_run(
    " intervenciones y acciones pedagógicas en el Año 2021, en el marco de la \
    Norma Técnica “Disposiciones para la implementación de las intervenciones \
    y acciones pedagógicas del Ministerio de Educación en los Gobiernos Regionales \
    y Lima Metropolitana en el Año Fiscal 2021”, aprobada mediante \
    RM N° 043-2021-MINEDU y modificada RM N° 159-2021-MINEDU." )
    interv_parrafo1.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
    tabla_interv = document.add_table(tabla1_region.shape[0]+1, tabla1_region.shape[1])
    tabla_interv.style = "Colorful List Accent 1"
    for j in range(tabla1_region.shape[-1]):
        tabla_interv.cell(0,j).text = tabla1_region.columns[j]
    for i in range(tabla1_region.shape[0]):
        for j in range(tabla1_region.shape[-1]):
            tabla_interv.cell(i+1,j).text = str(tabla1_region.values[i,j])
    interv_parrafo2 = document.add_paragraph(
    "A través de los Decretos Supremos N°s 092, 169, 189, 209 y 210-2021-EF, \
    se realizaron todas las transferencias de partidas programadas para el Año \
    Fiscal 2021 para el financiamiento de las intervenciones y acciones pedagógicas \
    hasta el 31 de diciembre.", style="List Bullet")
    interv_parrafo2.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
    interv_parrafo3 = document.add_paragraph(
    "Es importante considerar que la ejecución en los Contratos Administrativos \
    de Servicios (CAS) se ha visto afectada por la vigencia de la Ley N° 31131. \
    Por otro lado, la ejecución en bienes y servicios (excluyendo CAS) es menor \
    a lo esperado debido al bajo número de IIEE que brindan servicios presenciales \
    y semipresenciales a nivel nacional.", style="List Bullet")
    interv_parrafo3.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
    interv_parrafo4 = document.add_paragraph(
    "Las Unidades Ejecutoras de Educación de la región " , style="List Bullet")
    interv_parrafo4.add_run(region)
    interv_parrafo4.add_run(" cuentan con ")
    interv_parrafo4.add_run(pim_intervenciones_region)
    interv_parrafo4.add_run(
    " millones en su Presupuesto Institucional Modificado (PIM) para el \
    financiamiento de intervenciones y acciones pedagógicas, de los cuales se han ejecutado S/ ")
    interv_parrafo4.add_run(ejecucion_intervenciones_region)
    interv_parrafo4.add_run(".")
    interv_parrafo4.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
    #############################################################
    # Incluimos sección 2 Mascarillas y protectores faciales
    document.add_heading("2. Mascarillas y protectores faciales", level=1)
    mascarillas_parrafo1 = document.add_paragraph(
    "Mediante el Decreto de Urgencia N° 021-2021 y la Resolución de Secretaría \
General N° 047-2021-MINEDU, se transfirieron S/ ", style="List Bullet")
    mascarillas_parrafo1.add_run(transferencia_mascarilla)
    mascarillas_parrafo1.add_run(" millones de soles para \
la adquisición y distribución de mascarillas faciales textiles de uso \
comunitario para estudiantes y personal que labora en instituciones \
educativas públicas, así como protectores faciales para el mencionado \
personal.")
    mascarillas_parrafo1.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
    mascarillas_parrafo2 = document.add_paragraph(
    "La adquisición de mascarillas y protectores faciales es condición necesaria \
para el retorno seguro a los servicios educativos presenciales y \
semipresenciales, según lo dispuesto por las “Disposiciones para la \
prestación del servicio en las instituciones y programas educativos \
públicos y privados de la Educación Básica de los ámbitos urbanos y \
rurales, en el marco de la emergencia sanitaria de la COVID-19”, \
aprobado mediante Resolución Ministerial N° 121-2021- MINEDU y modificado \
con Resoluciones Ministeriales N° 199-2021-MINEDU y N° 273-2021- MINEDU.", style="List Bullet")
    mascarillas_parrafo2.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
    mascarillas_parrafo3 = document.add_paragraph(
    "Con fecha de corte al ", style="List Bullet")
    mascarillas_parrafo3.add_run(fecha_corte_mascarillas)
    mascarillas_parrafo3.add_run(
    ", la ejecución a nivel regional de los recursos de mascarillas faciales \
textiles protectores faciales fue del ")
    mascarillas_parrafo3.add_run(devengado_mascarillas)
    mascarillas_parrafo3.add_run(" (devengado) según se presenta a continuación:")
    mascarillas_parrafo3.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY
    tabla2_interv = document.add_table(tabla2_region.shape[0]+1, tabla2_region.shape[1])
    tabla2_interv.style = "Colorful List Accent 1"
    for j in range(tabla2_region.shape[-1]):
        tabla2_interv.cell(0,j).text = tabla2_region.columns[j]
    for i in range(tabla2_region.shape[0]):
        for j in range(tabla2_region.shape[-1]):
            tabla2_interv.cell(i+1,j).text = str(tabla2_region.values[i,j])
    document.save(here() / f'output/{region}_AM.docx')


# # Contenido -------------------------------------------------------------------

# # Creación del documento en base al template
# document = docx.Document(here() / "input/formato.docx")

# #Título del documento
# title=document.add_heading('AYUDA MEMORIA', 0)
# run = title.add_run()
# run.add_break()
# title.add_run('REGIÓN ')
# title.add_run(region)
# run = title.add_run()
# run.add_break()
# title.add_run(datetime.today().strftime('%d-%m-%y'))


# # 1) Intervenciones pedagógicas

# document.add_heading("1. Intervenciones pedagógicas", level=1)

# interv_parrafo1 = document.add_paragraph(
#     "Las Unidades Ejecutoras de Educación de la región" , style="List Bullet")
# interv_parrafo1.add_run(region)
# interv_parrafo1.add_run(" vienen implementando ")
# interv_parrafo1.add_run(numero_intervenciones)
# interv_parrafo1.add_run(
#     " intervenciones y acciones pedagógicas en el Año 2021, en el marco de la \
# Norma Técnica “Disposiciones para la implementación de las intervenciones \
# y acciones pedagógicas del Ministerio de Educación en los Gobiernos Regionales \
# y Lima Metropolitana en el Año Fiscal 2021”, aprobada mediante \
# RM N° 043-2021-MINEDU y modificada RM N° 159-2021-MINEDU." )
# interv_parrafo1.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

# interv_parrafo2 = document.add_paragraph(
#     "A través de los Decretos Supremos N°s 092, 169, 189, 209 y 210-2021-EF, \
# se realizaron todas las transferencias de partidas programadas para el Año \
# Fiscal 2021 para el financiamiento de las intervenciones y acciones pedagógicas \
# hasta el 31 de diciembre.", style="List Bullet")
# interv_parrafo2.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

# interv_parrafo3 = document.add_paragraph(
#     "Es importante considerar que la ejecución en los Contratos Administrativos \
# de Servicios (CAS) se ha visto afectada por la vigencia de la Ley N° 31131. \
# Por otro lado, la ejecución en bienes y servicios (excluyendo CAS) es menor \
# a lo esperado debido al bajo número de IIEE que brindan servicios presenciales \
# y semipresenciales a nivel nacional.", style="List Bullet")
# interv_parrafo3.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

# interv_parrafo4 = document.add_paragraph(
#     "Las Unidades Ejecutoras de Educación de la región " , style="List Bullet")
# interv_parrafo4.add_run(region)
# interv_parrafo4.add_run(" cuentan con ")
# interv_parrafo4.add_run(pim_intervenciones)
# interv_parrafo4.add_run(
#     " millones en su Presupuesto Institucional Modificado (PIM) para el \
# financiamiento de intervenciones y acciones pedagógicas, de los cuales se han ejecutado S/ ")
# interv_parrafo4.add_run(ejecutado_intervenciones)
# interv_parrafo4.add_run(".")
# interv_parrafo4.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY



# # 2) Mascarillas y protectores faciales

# document.add_heading("2. Mascarillas y protectores faciales", level=1)

# mascarillas_parrafo1 = document.add_paragraph(
#     "Mediante el Decreto de Urgencia Nº 021-2021 y la Resolución de Secretaría \
# General N° 047-2021-MINEDU, se transfirieron S/ 96.3 millones de soles para \
# la adquisición y distribución de mascarillas faciales textiles de uso  \
# comunitario para estudiantes y personal que labora en instituciones \
# educativas públicas, así como protectores faciales para el mencionado \
# personal.", style="List Bullet")
# mascarillas_parrafo1.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

# mascarillas_parrafo2 = document.add_paragraph(
#     "La adquisición de mascarillas y protectores faciales es condición necesaria \
# para el retorno seguro a los servicios educativos presenciales y \
# semipresenciales, según lo dispuesto por las “Disposiciones para la \
# prestación del servicio en las instituciones y programas educativos \
# públicos y privados de la Educación Básica de los ámbitos urbanos y \
# rurales, en el marco de la emergencia sanitaria de la COVID-19”, \
# aprobado mediante Resolución Ministerial N° 121-2021- MINEDU y modificado \
# con Resoluciones Ministeriales N° 199-2021-MINEDU y N°273-2021- MINEDU.", style="List Bullet")
# mascarillas_parrafo2.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

# mascarillas_parrafo3 = document.add_paragraph(
#     "Con fecha de corte al ", style="List Bullet")
# mascarillas_parrafo3.add_run(fecha_corte_mascarillas)
# mascarillas_parrafo3.add_run(
#     " la ejecución a nivel regional de los recursos de mascarillas faciales \
#     textiles protectores faciales fue del ")
# mascarillas_parrafo3.add_run(devengado_mascarillas)
# mascarillas_parrafo3.add_run(" (devengado) según se presenta a continuación:")
# mascarillas_parrafo3.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

# # 3) Compromisos de Desempeño

# document.add_heading("3. Compromisos de Desempeño", level=1)

# compromisos_parrafo1 = document.add_paragraph(
#     "En el marco de la Norma Técnica para la implementación del mecanismo \
# denominado Compromisos de Desempeño 2021, aprobada por Resolución \
# Ministerial N° 042-2021-MINEDU y modificada por la Resolución Ministerial \
# N° 160-2021-MINEDU, se han realizado transferencias de partidas a favor \
# de las Unidades Ejecutoras de Educación del Gobierno Regional de ", style="List Bullet")
# compromisos_parrafo1.add_run(region)
# compromisos_parrafo1.add_run(" por la suma total de ")
# compromisos_parrafo1.add_run(transferido_compromisos)
# compromisos_parrafo1.add_run(" De dichos recursos, ")
# compromisos_parrafo1.add_run(acciones_centrales_cdd)
# compromisos_parrafo1.add_run(" corresponden a las acciones centrales, según el \
#                              siguiente detalle")
# compromisos_parrafo1.paragraph_format.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY


# # Guardar output
# document.save(here() / "output/AM.docx")