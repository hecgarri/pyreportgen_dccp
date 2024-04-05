# -*- coding: utf-8 -*-
"""
Titulo: Automatizacion reportes regionales (V3)
Created on Fri Dic 2 16:24:52 2023

Proceso que automatiza reportes regionales, generando un documento word por cada region a partir de template y Base de datos.

Obs: Se deben revisar los inputs antes de correr
     Refinar funciones
     Eliminar librerías sin uso

@author: diego.martinez
"""

# =============================================================================
# Extracción en python de datos de comunicado de compras a nivel regional
# =============================================================================

#Seteo de librerias

import pandas as pd
import numpy as np
import sqlalchemy as sa              #Para conexión a BD, requerido para usar pd.read_sql()
import urllib                        #Para formatear string de conexión

import docx
import matplotlib.pyplot as plt 
import matplotlib.patches as mpatches
import openpyxl as opxl             # OJO EN LA CASA para importar excel como dataframe

from docx import Document
from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage   #rellenar templates

import os
from pylab import savefig
import itertools

import pyodbc                               ### another engine to make DB connection and run the queries. Alternative to sqlalchemy
from itertools import repeat
import querysGR as qg

import sys
import define_funciones as dfn

#Transforma codigo a word
#Agregar github


#Conexion a DW

#param = urllib.parse.quote_plus("DRIVER={SQL Server};SERVER=10.34.71.202;UID=datawarehouse;PWD=datawarehouse;DATABASE=DM_Transaccional")#;Trusted_Connection=yes
#conn = sa.create_engine("mssql+pyodbc:///?odbc_connect=%s" % param)

#param = urllib.parse.quote_plus("DRIVER={SQL Server};SERVER=10.34.71.227;UID=eymetric;PWD=Eym3tr1c.2022;DATABASE=DM_Transaccional_2022")#;Trusted_Connection=yes
#conn = sa.create_engine("mssql+pyodbc:///?odbc_connect=%s" % param)

#A DCCPProcurement
#paramAQ = urllib.parse.quote_plus("DRIVER={ODBC Driver 13 for SQL Server};DATABASE=DCCPProcurement;SERVER=10.34.71.145;UID=datawarehouse;PWD=datawarehouse")
#bd = sa.create_engine("mssql+pyodbc:///?odbc_connect=%s" % paramAQ)
#paramAQ = urllib.parse.quote_plus("DRIVER={ODBC Driver 13 for SQL Server};SERVER=10.34.71.146;UID=datawarehouse;PWD=datawarehouse")
#bd = sa.create_engine("mssql+pyodbc:///?odbc_connect=%s" % paramAQ)


#**** BBDD *****

# Verificar Drivers

###### DW ######


param_DW = urllib.parse.quote_plus("DRIVER={SQL Server};SERVER=10.34.71.202;UID=datawarehouse;PWD=datawarehouse;DATABASE=DM_Transaccional;TrustServerCertificate=yes")
conn_DW = sa.create_engine("mssql+pyodbc:///?odbc_connect=%s" % param_DW)

##### DEV #####

#/opt/microsoft/msodbcsql18/lib64/libmsodbcsql-18.2.so.2.1

#param_dev = 'mssql+pyodbc://datawarehouse:datawarehouse@10.34.71.227:1433/DM_Transaccional_dev?driver=ODBC+Driver+17+for+SQL+Server&TrustServerCertificate=yes'
#conn_dev = sa.create_engine(param_dev)

string_datawarehouse='mssql+pyodbc://datawarehouse:datawarehouse@10.34.71.227:1433/DM_Transaccional_dev?driver=ODBC+Driver+17+for+SQL+Server&TrustServerCertificate=yes'

engine = sa.create_engine(string_datawarehouse, fast_executemany=True)
conn_dev = engine.connect()

##### AQUILES #####

param_AQ = urllib.parse.quote_plus("DRIVER={SQL Server};DATABASE=DCCPProcurement;SERVER=10.34.71.146;UID=datawarehouse;PWD=datawarehouse;DATABASE=DCCPProcurement;TrustServerCertificate=yes")
conn_AQ = sa.create_engine("mssql+pyodbc:///?odbc_connect=%s" % param_AQ)

### connections with pyodbc ###
conn_AQ_pyodbc=pyodbc.connect('DRIVER={SQL Server};DATABASE=DCCPProcurement;SERVER=10.34.71.146;UID=datawarehouse;PWD=datawarehouse;DATABASE=DCCPProcurement;TrustServerCertificate=yes')#;Encrypt=yes')#;TrustServerCertificate=yes')#;TrustServerCertificate=yes')
#conn_AQ_pyodbc = pyodbc.connect('DRIVER={SQL Server};SERVER=10.34.71.146;DATABASE=DCCPProcurement;UID=datawarehouse;PWD=datawarehouse')
cursor=conn_AQ_pyodbc.cursor()
#*************************

# =============================================================================
# Setteo Diccionarios Regionales y Nacional
# =============================================================================
                                  
# Diccionario Regional, adaptar nombres de BBDD
# Usar BBDD? Renombrar para diccionario con mas datos??
RegNomb =  {'Antofagasta':              {'nom':'Región de Antofagasta',                  'nomCt':'Antofagasta'},
            'Araucanía':                {'nom':'Región de La Araucanía',                 'nomCt':'La Araucanía'},
            'Arica y Parinacota':       {'nom':'Región de Arica y Parinacota',           'nomCt':'Arica y Parinacota'},
            'Atacama':                  {'nom':'Región de Atacama',                      'nomCt':'Atacama'},
            'Aysén':                    {'nom':'Región de Aysén',                        'nomCt':'Aysén'},
            'Bío-Bío':                  {'nom':'Región del Bío-Bío',                     'nomCt':'el Bío-Bío'},
            'Coquimbo':                 {'nom':'Región de Coquimbo',                     'nomCt':'Coquimbo'},
            "Lib. Gral. Bdo. O'Higgins":{'nom':"Región de O'Higgins",                    'nomCt':"O'Higgins"},
            'Los Lagos':                {'nom':'Región de Los Lagos',                    'nomCt':'Los Lagos'},
            'Los Ríos':                 {'nom':'Región de Los Ríos',                     'nomCt':'Los Ríos'},
            'Magallanes y Antártica':   {'nom':'Región de Magallanes y la Antártica',    'nomCt':'Magallanes'},
            'Maule':                    {'nom':'Región del Maule',                       'nomCt':'el Maule'},
            'Metropolitana':            {'nom':'Región Metropolitana',                   'nomCt':'la R. Metropolitana'},
            'Ñuble':                    {'nom':'Región del Ñuble',                       'nomCt':'el Ñuble'},
            'Tarapacá':                 {'nom':'Región de Tarapacá',                     'nomCt':'Tarapacá'},
            'Valparaíso':               {'nom':'Región de Valparaíso',                   'nomCt':'Valparaíso'}}

#nombres meses para usar con mes_i y mes_f (int)
meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 
         'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']

###########################################################
###    Def parámetros indicadores y ejecución querys    ###
###########################################################

#juntar parámetros al principio, querys al final. agregar condicionalidad a ejecución de querys

#Años y meses
anoReg = 2023
anoRegM = anoReg - 1
mes_i = 1
mes_f = 12


#Setea lista de regiones
print('Ejecutando consulta para ver listado de regiones')

listReg = pd.read_sql(con = conn_dev,  sql = qg.queryRegiones() )

listReg = list(listReg['Region'])

print('El resultado de la consulta es:')
print(listReg)



#Totales Regionales 2023 ¿y 2022?
listColTmnReg = ['Tmn']
dtoTmnReg = 'tmnReg'


#Parametros Top Modalidad de compras
listColModReg = ['Mod']
dtoModReg = 'modReg'

#Query reg pa instituciones
listColInsReg = ['Ins']
dtoInsReg = 'insReg'


#Proveedores
dtoPrvReg = 'prvReg'
listColPrvReg = ['Prv','PrvID']


#Montos Rubros Regionales
topRubReg = 5 #query de rubro requiere el top
                #como se usarán datos solo para el gráfico se decide que el  top 5 es una cantidad adecuada
listColRubReg = ['Rub', 'Rank']
dtoRubReg = 'rubReg'


#Montos OC Regionales
listColOCReg = ['OCod', 'Ins', 'Prv', 'PrvID', 'Mtv', 'OLink']
dtoOCReg = 'ocReg'


#Montos Sectores Regionales
listColSecReg = ['Sec']
dtoSecReg = 'secReg'

#---querys

#Setea lista de regiones
listReg = pd.read_sql(con = conn_dev,  sql = qg.queryRegiones() )
listReg = list(listReg['Region'])



#Totales Regionales 2023 ¿y 2022?
print('Ejecutando consulta sobre totales regionales:')
TotReg = pd.read_sql(con = conn_dev,  sql = qg.queryTotalRegion(mes_i, mes_f))
print('El resultado de la consulta es: ')
print(TotReg)

print('Ejecutando consulta sobre totales regionales')
TotRegBig = pd.read_sql(con = conn_dev,  sql = qg.theQueryReg(mes_i, mes_f)) #Reemplazar al de arriba?? demora
# se puede optimizar el tiempo de carga eliminando los requerimientos de para OC en theQueryReg()
#           opcionalmente se puede usar queryOrdenCompraRegionTop() para consultar las OC
print('El resultado de la consulta es:')
#Montos Rubros Regionales
RubReg = pd.read_sql(con = conn_dev,  sql = qg.queryRubroRegion(mes_i, mes_f, topRubReg) )
print(RubReg)
#Montos Compra Agils Regional 2022 y 2023
print('Ejecutando consulta sobre la compra ágil')
CAReg = pd.read_sql(con = conn_dev,  sql = qg.queryCompraAgilRegion(mes_i, mes_f) )
print(CAReg)

########################################
# Grafico Regional Sectores por 3 años #
########################################


listaAnos = [2023, 2022, 2021]
print('Ejecutando consulta sobre sectores')
listSectores = pd.read_sql(con = conn_dev,  sql = qg.querySectores())
listSectores = listSectores['Sec']
print('El resultado de la consulta es:')
print(listSectores)

#Cambiar nombre de sector Gob. Central, Universidades a modo abreviado
listSectores2 = listSectores
#listSectores
#idx_sec = np.where(listSectores2=='Gob. Central, Universidades')[0][0]
#listSectores2[idx_sec] = 'Gob. Central, Ues.'

print('Ejecutando consulta sobre sectores en las regiones')
dfSecRegAnos = []
for i in range (0, len(listaAnos)):
    dfSecRegAnos.append(pd.read_sql(con = conn_dev,  sql = qg.querySectorRegion(mes_i, mes_f, listaAnos[i])))
dfSecRegAnos = pd.concat(dfSecRegAnos)

print(dfSecRegAnos)    

#exporta datos a excel (útil para gráficos de montos de sector por región, montos de sector por región por años, montos de sector por región nacional)
dfSecRegAnos.to_excel('SectoresRegionesAños.xlsx')


########################## Guardar dataframes de la sesión #######################################
#se eliminan las conexiones ya que no se pueden guaradar
#del([param_DW,param_dev,conn_dev,param_AQ,conn_AQ,conn_AQ_pyodbc,cursor,conn_DW])

#import dill                            #pip install dill --user
#filename = 'Reportes2023_procesados_20240129.pkl'
#dill.dump_session(filename)

# and to load the session again:
#dill.load_session(filename)

###################################################
###   Llenado diccionario/contexto y template   ###
###################################################

#Importar template
docu = DocxTemplate("docxtpl.docx")

#Settear diccionario nacional
contextoNac = {'ano' : str(anoReg),
               'anoM' : str(anoRegM),
               'mesI' : meses[mes_i - 1],   #llama nombre del mes para usar en reporte
               'mesF' : meses[mes_f - 1]}

#settear diccionario adicionales regionales
adicionalesReg = dfn.impAdicionalesReg()



#prepara dataframe para gráficos de sector por región por año (se crea y agrega en iterador de regiones)
    #y sector por región nacional (se crea y agrega antes de iterador de regiones)

#filta año actual para gráfico sector por región nacional
dfSectorReg = dfSecRegAnos.loc[dfSecRegAnos['Ano'] == 2023]
print(dfSectorReg)
print(listReg)
print(listSectores)
#crea  diccionario para gráfico sector por región nacional
dcSectorReg = dfn.diccionarioGraficos(dfSectorReg, listReg, 'Region', listSectores, 'Sec')



#Crea gráfico
tituloGrafSecRegNac = 'Sector por región, nivel nacional'
dfn.graficoSectoresRegion(dcSectorReg, listSectores, tituloGrafSecRegNac)

#agrega gráfico al diccionario
img = InlineImage(docu, tituloGrafSecRegNac+'.png',width=Inches(7))
contextoNac.update({'secRegNacGrf' : img})



#diccionario que almacena los diccionarios de cada región para luego exportar una planilla con todos los datos
dicGlob = {}

anoReg = 2023
anoRegM = anoReg - 1

#Iterador de regiones/documentos
for r in listReg:
  
   print('##############')
   print(r)
   print(anoReg)
   print('##############')

   contexto = dfn.setContextoReg(r, RegNomb)

   rTotReg = TotReg.loc[TotReg['Region'] == r]
   print(type(rTotReg))

   contexto.update(dfn.agregarTotalesRegion(rTotReg, anoReg))
   
   #tmn Reg
   rTmnReg = dfn.extraerDataframe (TotRegBig, r, listColTmnReg)
   rTmnReg = dfn.fmtoDataframe(rTmnReg, listColTmnReg)
   contexto.update(dfn.dataframeDiciconario(rTmnReg, dtoTmnReg))

    #modalidad de compra
   rModReg = dfn.extraerDataframe (TotRegBig, r, listColModReg)
   rModReg = dfn.fmtoDataframe(rModReg, listColModReg)    
   contexto.update(dfn.dataframeDiciconario(rModReg, dtoModReg))

    
    # proveedores a los que se le compra por región
   rPrvReg = dfn.extraerDataframe (TotRegBig, r, listColPrvReg)
   rPrvReg = dfn.fmtoDataframe(rPrvReg, listColPrvReg) 
   contexto.update(dfn.dataframeDiciconario(rPrvReg, dtoPrvReg))

    # top Instituciones reg
   rInsReg = dfn.extraerDataframe (TotRegBig, r, listColInsReg)
   rInsReg = dfn.fmtoDataframe(rInsReg, listColInsReg)    
   contexto.update(dfn.dataframeDiciconario(rInsReg, dtoInsReg))

    #top rubro por region (usa dataframe personalizado)
    
   RubReg.to_excel('rubroRegion.xlsx')
   rRubReg = dfn.extraerDataframe(RubReg, r, listColRubReg)
   rRubReg = dfn.fmtoDataframe(rRubReg, listColRubReg)
    
   ttlGrafRub = 'Rubros más transados, '+ RegNomb[r]['nomCt']+' (Millones de Pesos)'
   dfn.grafBarras((np.array(rRubReg['CLP'])/1000000), rRubReg[''], ttlGrafRub)
   img = InlineImage(docu, ttlGrafRub+'.png',width=Inches(7))
   contexto.update({'rubRegGrf' : img})

    #no lo pide ahora comunicaciones
    #contexto.update(dataframeDiciconario(rRubReg, dtoRubReg, topRubReg))
    
   rOCReg = dfn.extraerDataframe(TotRegBig, r, listColOCReg)
   rOCReg = dfn.fmtoDataframe(rOCReg, listColOCReg)
   contexto.update(dfn.dataframeDiciconario(rOCReg, dtoOCReg))

   rSecReg = dfn.extraerDataframe(TotRegBig, r, listColSecReg)
   rSecReg = dfn.fmtoDataframe(rSecReg, listColSecReg)
   contexto.update(dfn.dataframeDiciconario(rSecReg, dtoSecReg))

   ttlGrafSec = 'Porcentaje participación por Sector, '+ RegNomb[r]['nomCt']
   dfn.grafTorta(rSecReg['CLP'], rSecReg[''], ttlGrafSec)
   img = InlineImage(docu, ttlGrafSec+'.png',width=Inches(7))
   contexto.update({'secRegGrf' : img})

   dfSectorAno = dfSecRegAnos.loc[dfSecRegAnos['Region'] == r]
   dcSectorAno = dfn.diccionarioGraficos(dfSectorAno, listaAnos, 'Ano', listSectores, 'Sec')
   tituloGrafBarra3 = 'Monto por sector cada año, '+ RegNomb[r]['nomCt']
   dfn.grafBarrasTriple(dcSectorAno, listaAnos, listSectores, tituloGrafBarra3)
   img = InlineImage(docu, tituloGrafBarra3+'.png',width=Inches(7))
   contexto.update({'secRegAnoGrf' : img})

   rCAReg = CAReg.loc[CAReg['Region'] == r]
   contexto.update(dfn.agregarCARegion(rCAReg))

   contexto.update(adicionalesReg[r])

   contexto.update(contextoNac)
    
    
   #print(contexto)
   dicGlob.update({r: contexto})
   print(dicGlob)
   docu.render(contexto)
    #os.remove(titGrafSecReg +'.png')    #según cantidad de gráficos se podría iterar
    
   nomDocu = contexto['ano'] + ' cifras regionales ' + contexto['nomCt'] + ' ' + contexto['mesI'] + '-' + contexto['mesF']
   docu.save('reportes/'+nomDocu+'.docx')


#######################################################
# Exporta planilla con todos los contextos regionales #
#######################################################

dfGlob = pd.DataFrame(dicGlob)
dfGlob.to_excel('dfGlob.xlsx')

#############################
# Borar gráficos de carptea #
#############################


os.remove('Sector por región, nivel nacional'+'.png')
for r in listReg:
    
    os.remove('Monto por sector cada año, '+ RegNomb[r]['nomCt']+'.png')
    os.remove('Rubros más transados, '+ RegNomb[r]['nomCt']+' (Millones de Pesos)' +'.png')   
    os.remove('Porcentaje participación por Sector, '+ RegNomb[r]['nomCt'] +'.png')   
