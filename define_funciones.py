
###########################################
###########################################
######                               ######
######       Definir Funciones       ######
######                               ######
###########################################
###########################################


"""
Módulo para el análisis y visualización de datos financieros y de compras públicas.

Este módulo proporciona una serie de funciones útiles para el análisis y visualización de datos de compras públicas. Incluye funciones para formatear números, generar gráficos de barras, gráficos de sectores, gráficos de torta, y más. Además, ofrece herramientas para trabajar con dataframes, calcular totales y porcentajes, y agregar datos adicionales por región.

Funciones:
- fmtoEntero: Formatea un número con puntos y símbolos monetarios.
- fmtoPorcien: Convierte un valor en porcentaje y lo formatea como string.
- palabraVar: Retorna una cadena que describe si una variación es un aumento, disminución o no variación.
- diccionarioGraficos: Transforma datos de un dataframe en un diccionario de listas para graficar.
- grafBarrasTriple: Genera un gráfico de barras agrupadas de a 3.
- graficoSectoresRegion: Genera un gráfico de sectores por región.
- grafTorta: Genera un gráfico de torta.
- grafBarras: Genera un gráfico de barras horizontal.
- setContextoReg: Configura un diccionario de contexto regional.
- agregarTotalesRegion: Calcula y agrega totales de monto y OC por región.
- agregarGrafMontoSectorRegion: Agrega gráficos de monto por sector y región.
- agregarCARegion: Agrega datos de la CA por región.
- extraerDataframe: Filtra un dataframe por región y columnas especificadas.
- fmtoDataframe: Formatea un dataframe para su visualización.
- dataframeDiciconario: Transforma un dataframe filtrado en un diccionario.
- impAdicionalesReg: Importa datos adicionales para cada región desde un archivo Excel.

Autor: [Tu nombre]
Fecha de creación: [Fecha de creación del módulo]
Versión: 1.0
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt 
import matplotlib.patches as mpatches
import openpyxl as opxl 

from docx import Document
from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage   #rellenar templates




#entra numero y retorna valor con puntos en string
#abrevia en millones si supera 8 dígitos
#agrega símbolos monetarios
def fmtoEntero(x, mnd=''):
    ini = ''
    fin = ''

    #abrevia si supera 8 digitos
    if x >= 100000000: #me pareció que muestre al menos 3 dígitos
        x = x / 1000000
        fin = ' millones'
    
    #verificar formato de escritura
    if mnd == 'CLP':
        ini = '$'
    elif mnd == 'USD':
        ini = 'US $'
        #fin = fin + ' USD'
    elif mnd == 'CLF':
        if fin == ' millones':
            fin = fin + ' de UF'
        else:
            fin = fin + ' UF'
    
    x = format(int(round(x)),',d') #crea str de entero con puntuación
    x = x.replace(",",".") #formato de puntos latino

    return ini + x + fin


#entra tasa y retorna el porcentaje como string
def fmtoPorcien(x):
    x = x * 100
    x = format(x,',.1f') #crea str con 1 decimal y puntuación
    x = x.replace(".","a").replace(",",".").replace("a",",") #formato de puntos latino
    return x+'%'


#retorna string aumento/disminución según variacón numérica entregada
def palabraVar(var):
    plb = 'una no variación' #es necesario?
    if var > 0:
        plb = 'un aumento'
    elif var < 0:
        plb = 'una disminución'
    return plb



#Transforma ordenadamente datos de un dataframe a un diccionario de listas, ordenando los valores de la columna 'CLP' para graficarlos
    #Para ordenarlos usa lstArrays como primer criterio para filtrar el dataframe y para nombrar las claves de las listas del diccionario
    #Usa listaDatos como segundo criterio para filtrar dataframe y ordenar los datos dentro de las listas
    #lstaArrays y lstaDatos deben contener datos de existentes de columnas del dataframe cada uno, nombres especificados con las variabels colArrays y colDatos
    #Ej: Si tenemos el dataframe con columnas 'Año', 'Sector' y 'CLP', usamos lstaArrays [2023,  2022] con 'Año', listaDato ['FFAA', 'Salud'] con 'Sector'
    #Retornaría el siguiente diccionario:
    #{2023 : [(monto FFAA en 2023), (monto Salud en 2023)],
    # 2022 : [(monto FFAA en 2022), (monto Salud en 2022)]}
    #
    #(df = dataframe, lstaArrays = lista, colArrays = string, lstaDatos = lista, colDatos = strin)
def diccionarioGraficos(df, lstaArrays, colArrays, lstaDatos, colDatos):

    #Crea diccionario a retornar
    dic={}

    #Recorre la lstaArrays
    for i in range (0, len(lstaArrays)):

        #Agrega al diccionario elemento con el dato de lstaArrays actual como clave y una lista vacía como valor
        dic[lstaArrays[i]]=[]
        print('')
        print('Año '+str(lstaArrays[i]))

        #Recorre la lstaDatos
        for j in range (0, len(lstaDatos)):

            #Extrae del dataframe monto en CLP filtrando filas por el valor de lstaArrays y lstaDatos actuales
            #Ej: filtro 'Año' ==  2022 y 'Sector' == FFAA
            clp = df.loc[(df[colArrays]==lstaArrays[i]) & (df[colDatos]==lstaDatos[j]), 'CLP'].values[0]
            #Agrega monto a la lista de clave correspondiente en lstaArray y en la posicción correspondiente de lstaDatos
            dic[lstaArrays[i]].append(clp)

        print(dic)

    return dic



#Genera gráfico de barras agrupadas de a 3
#Requiere dic (un diccionario de listas), lstaArrays (una lista del largo de dic y que contiene las claves de dic),
        #lstaDatos (una lista del largo de las listas de dic, contiene la etiqueta y orden de los datos de las listas de dic)
        #y titulo (string que da nómbre al gráfico)
#puede modificarse para que la cantidad de de barras por etiqueta se adapte a el largo de lstaArrays
def grafBarrasTriple(dic, lstaArrays, lstaDatos, titulo):
    plt.clf()

    # set width of bar 
    barWidth = 0.25
    fig = plt.subplots(figsize =(12, 8)) 

    # Set position of bar on X axis 
    br1 = np.arange(len(dic[2023])) 
    br2 = [x + barWidth for x in br1] 
    br3 = [x + barWidth for x in br2] 
    
    # Make the plot
    plt.bar(br1, dic[lstaArrays[0]], 
            width = barWidth, 
            label =str(lstaArrays[0])) 
    plt.bar(br2, dic[lstaArrays[1]], 
            width = barWidth, 
            label =str(lstaArrays[1])) 
    plt.bar(br3, dic[lstaArrays[2]], 
            width = barWidth, 
            label =str(lstaArrays[2])) 
    
    # Adding Xticks 
    plt.xlabel(titulo, fontweight ='bold', fontsize = 15) 
    plt.ylabel('Millones de Pesos', fontweight ='bold', fontsize = 15) 
    plt.xticks([r + barWidth for r in range(len(dic[2023]))], 
            lstaDatos)
    
    plt.legend()
    plt.savefig(titulo+'.png', bbox_inches = 'tight')



#Necesita dic (diccionario) de listas, las claves indican la region, lstaDatos (lista) indica los sectores y orden de los datos en las listas de dic
def graficoSectoresRegion(dic, lstaDatos, titulo):
    plt.clf()

    #Transforma los datos de cada lista a su valor en porcentaje sobre el total de cada lista
    for key, value in dic.items():
        total = np.sum(value)
        dic[key] = np.round(np.array(value)*100/total, 1)

    """
    Parameters
    ----------
    results : dict
        A mapping from question labels to a list of answers per category.
        It is assumed all lists contain the same number of entries and that
        it matches the length of *category_names*.
    category_names : list of str
        The category labels.
    """
    labels = list(dic.keys())
    data = np.array(list(dic.values()))
    data_cum = data.cumsum(axis=1)
    category_colors = plt.colormaps['turbo'](
        np.linspace(0.15, 0.85, data.shape[1]))

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.invert_yaxis()
    ax.xaxis.set_visible(False)
    ax.set_xlim(0, np.sum(data, axis=1).max())

    for i, (colname, color) in enumerate(zip(lstaDatos, category_colors)):
        widths = data[:, i]
        starts = data_cum[:, i] - widths
        rects = ax.barh(labels, widths, left=starts, height=0.5,
                        label=colname, color=color)

        r, g, b, _ = color
        text_color = 'white' if r * g * b < 0.5 else 'darkgrey'
        ax.bar_label(rects, label_type='center', color=text_color)
    ax.legend(ncols=len(lstaDatos), bbox_to_anchor=(0, 1),
              loc='lower left', fontsize='small')
    
    plt.savefig(titulo+'.png', bbox_inches = 'tight')
    return fig, ax






#Agrega gráfico de torta, necesita datos (x), etiqueta (labels) y título (titGraf)
def grafTorta(x, labels, titGraf): 
    plt.clf()
    
    fig, ax = plt.subplots(figsize =(10, 9))
    #plt.figure(figsize=(4.5,4.5))
    bordeG = {'linewidth' : 1, 'edgecolor' : 'white'}
    plt.pie(x
            , autopct='%.1f%%'
            , textprops=dict(color="grey", size=10)
            , pctdistance=1.15
            , wedgeprops = bordeG)
    legend = plt.legend(labels
                        , loc = "lower center"
                        , bbox_to_anchor=(0.27, -0.15, 0.5, 0.5)
                        , ncols = 3
                        , fontsize=10
                        , handlelength=0.7
                        , handleheight=0.7
                        ) 
    tituloG = plt.title(titGraf ,fontweight="bold",fontsize=16)

    
    plt.savefig(titGraf+'.png', bbox_inches = 'tight')



#Crea gráfico de barras horizontal, necesita datos (x), etiqueta (labels) y título (titulo)
def grafBarras(x, labels, titulo):
    plt.clf()
    fig, ax = plt.subplots(figsize =(12, 12))
    
    #color = ['lightblue', 'blue', 'purple', 'red', 'black']
    
    # Horizontal Bar Plot
    ax.barh(labels, x) #height=0.3, linewidth=0.1)
    #ax.legend(labels,loc="center", bbox_to_anchor=(1., 1.02) , borderaxespad=0., ncol=5)
    # Remove axes splines
    for s in ['top', 'right']:
        ax.spines[s].set_visible(False)
    
    # Remove x, y Ticks
    #ax.xaxis.set_ticks_position('none')
    #ax.yaxis.set_ticks_position('none')
    
    # Add padding between axes and labels
    ax.xaxis.set_tick_params(pad = 5)
    ax.yaxis.set_tick_params(pad = 10)
    
    # Add x, y gridlines
    ax.grid(#b = True, 
            color ='grey',
            linestyle ='-.', linewidth = 0.5,
            alpha = 0.2)
    
    # Show top values 
    ax.invert_yaxis()
    
    # Add annotation to bars
    for i in ax.patches:
        plt.text(i.get_width()+0.2, i.get_y()+0.5, 
                str(round((i.get_width()), 2)),
                fontsize = 15, fontweight ='bold',
                color ='grey')
    
    # Add Plot Title
    ax.set_title(titulo,
                loc ='left', 
                fontsize = 25)
    
    # Add Text watermark
    fig.text(0.9, 0.15, 'ChileCompra', fontsize = 12,
            color ='grey', ha ='right', va ='top',
            alpha = 0.7)
    plt.style.use('seaborn-v0_8-darkgrid')
    # Show Plot
    plt.savefig(titulo+'.png', bbox_inches = 'tight')



#Settea diccionario contexto regional, agrgando como primeros datos nombre largo y nombre corto de la región
def setContextoReg(r, regNom):
    ctxt = {}
    ctxt.update(regNom[r])
    return ctxt



#Retorna diccionario con total transado de monto y OC por región
#(dataframe filtrado por región, moneda a usar)
def agregarTotalesRegion(df, anoReg):
    anoRegM = anoReg-1
    nctxt = {}

    regMndUSD   = df.loc[df['Año'] == anoReg, 'Monto_Bruto_USD'].iloc[0]     
    regMnd      = df.loc[df['Año'] == anoReg, 'Monto_Bruto_'+'CLP'].iloc[0]     
    regOC       = df.loc[df['Año'] == anoReg, 'CantOC'].iloc[0]               
    regMndM     = df.loc[df['Año'] == anoRegM, 'Monto_Bruto_'+'CLP'].iloc[0]    
    regOCM      = df.loc[df['Año'] == anoRegM, 'CantOC'].iloc[0]              

    nctxt['totRegUSD']      = fmtoEntero( regMndUSD , 'USD')
    nctxt['totReg'+'CLP']     = fmtoEntero( regMnd    , 'CLP')
    nctxt['totRegOC']       = fmtoEntero( regOC )
    nctxt['totReg'+'CLP'+'M'] = fmtoEntero( regMndM   , 'CLP')
    nctxt['totRegOCM']      = fmtoEntero( regOCM    )
    
    tasaVar                 = (regMnd - regMndM) / regMndM
    nctxt['totRegPct']      = fmtoPorcien(tasaVar)
    nctxt['totRegVarPlb']   = palabraVar(tasaVar)

    tasaVarOC = (regOC-regOCM)/regOCM
    nctxt['totRegPctOC'] = fmtoPorcien(tasaVarOC)

    return nctxt



#Llama a creación de gráfico y retorna diccionario para agregarlo a template
#¿Pide moneda o dejamos default?
def agregarGrafMontoSectorRegion(df, titGraf, docu):
    
    grafTorta(df['CLP'], df[''], titGraf)
    #pedir docu como parametro?
    img = InlineImage(docu, titGraf+'.png',width=Inches(6))
    dctGrf = {'secRegGrf' : img}
    return dctGrf



def agregarCARegion(df): 
    nctxt = {}

    anoReg = 2023
    anoRegM = anoReg - 1
    
    # Dataframes año actual y anterior
    df0 = df.loc[df['Ano'] == anoReg] #MONTOCLP_CAg  MONTOUSD_CAg  CantOC_CAg
    dfM = df.loc[df['Ano'] == anoRegM]

    mto0    = df0['MONTOCLP_CAg'].iloc[0]
    oc0     = df0['CantOC_CAg'].iloc[0]
    mtoM    = dfM['MONTOCLP_CAg'].iloc[0]
    ocM     = dfM['CantOC_CAg'].iloc[0]
    mto0CLF = df0['MONTOCLF_CAg'].iloc[0]
    mtoMCLF = dfM['MONTOCLF_CAg'].iloc[0]

    mtoVar  = (mto0 - mtoM)/ mtoM   #Variacion monto
    ocDif   = oc0 - ocM             #Diferencia cantidad OC
    ocVar   = (oc0-ocM)/ocM         #Variación porcentual en la cantidad de órdenes de compra
    mtoVarR = (mto0CLF-mtoMCLF)/mto0CLF #Variación real en los montos de la Compra Ágil
    
    nctxt['caReg'+'CLP']       = fmtoEntero(mto0, 'CLP')
    nctxt['caRegOC']         = fmtoEntero(oc0)
    nctxt['caReg'+'CLP'+'M']   = fmtoEntero(mtoM, 'CLP')
    nctxt['caRegOCM']        = fmtoEntero(ocM)
    nctxt['caReg'+'CLP'+'Var'] = fmtoPorcien(mtoVar)
    nctxt['caRegOCDif']      = fmtoEntero(ocDif)
    nctxt['caRegOCpct']   = fmtoPorcien(ocVar) #Variación porcentual en cantidad de órdenes de compra
    nctxt['caReg'+'CLF'] = fmtoEntero(mto0CLF)
    nctxt['caReg'+'CLF'+'M'] = fmtoEntero(mtoMCLF)
    nctxt['caReg'+'CLF'+'Var'] = fmtoPorcien(mtoVarR)

    return nctxt



def extraerDataframe (df, r, listCol):
    df = df.loc[df['Region'] == r]
    listCol = ['Ano', 'Region'] + listCol
    
    df = df.groupby(listCol).agg({'USD' : sum,
                                  'CLP' : sum,
                                  'CLF' : sum,
                                  'OC'  : sum,})
    
    return df



def fmtoDataframe (df, listCol):
    df = df.reset_index()                               #reincorporar columnas del index
    df = df[df['Ano'] == 2023]                          #revisar cuando se usen mas años
    df = df.sort_values(by = 'USD', ascending = False)  #asegurar orden top
    df = df.reset_index()                               #reenumerar index
    df = df.rename(columns = {listCol[0] : ''})         #formatea nombre de columna con dato de interés
    #######################################
    # BORRAR COLUMNAS AÑO, REGION E INDEX #
    #######################################
    return df



#Función para transformar dataframe filtrado en diccionario.
    #df: Dataframe filtrado por región, año y columnas necesarias. Columna principal se reconoce  como '' (pd.df)
    #dto: nombre del dato de la sección que se trabaja. Ej: Rubro por región -> 'rubReg' (string)
    #top: cantidad de filas a procesar de las con mayores montos, por defecto 15. Se recomienda usar el 15 (int)
def dataframeDiciconario (df, dto, top=15):
    
    #Crea el diccionario a retornar
    nctxt = {}

    #Calcula  el monto total en la región. Se usa como referencia para porcentajes
    tot = df['CLP'].sum()
    #Agrega al diccionario el total en formato legible como dinero. Útil para verificar con totales de otros grupos de datos
    nctxt[dto+'TOTAL'] = fmtoEntero(tot, 'CLP') 

    #Comparala candtidad de filas disponibles vs la solicitada con 'top', deja la cantidad menor para iterar..
    if top > df.shape[0]:
        print('Cantidad de datos es menor a la requerida. ('+dto+')')
    top = min(top, df.shape[0]) #para que top no sea mayor a cantidad de datos

    #Filtra el df dejando los n valores mayores. n = top
    df = df.head(top)

    #Iterar el df celda por celda
    #Iterador de filtas. i (el index) indica el lugar en el top - 1 (parte desde el 0). row contiene todos los datos relacionados a la posición
    for i,row in df.iterrows():
        
        #Iterador de columnas. colu indica la columna o dato de la fila.
        for colu in df.columns:

            #Para explicar dto + colu + str(i1)
            #Se usa 'dto', 'colu' e 'i' para nombrar la clave del diccionario y que coincida con los parámetros de las variables en el template
                #Si estamos trabajando las modalidades de compra (dto = 'modReg')
                #Si estamos en la fila 3, la tercera modalidad con mas monto transado (i = 2)
                #Y se entra a la columna de cantidad OC involucradas (colu = 'OC')
                #La clave será 'modRegOC3', de cantidad de OC en la tercera modadlidad de compra de mayor monto transado
                #
                #Mismo caso, pero se entra a la columna que indica la modalidad. Al se la columna principal se identifica colu = ''
                #Clave 'modReg3', tercera modalidad de compra con mayor montro tranzado
            
            #El if reconoce si la celda es un número a formatear con fmtoEntero(), lo cual ocurre con las columnas 'USD', 'CLP', 'CLF', 'OC'
            if colu in ['USD', 'CLP', 'CLF', 'OC']:
                #Si cumple, guarda en el diccionario como lista
                nctxt[dto+colu+str(i+1)]     = [row[colu], colu]
            else:
                #Si no, guarda el valor normalmente
                nctxt[dto+colu+str(i+1)]     = row[colu]

        #Usando el total calculado al principio de la función (con CLP), se usa el valor CLP de la fila para calcular su porcentaje
            #respecto al total,  y lo argega al diciconario
            #Se hace fuera del iterador de columnas porque este valor no tiene columna calculada 
        nctxt[dto+'Pct'+str(i+1)]   = fmtoPorcien((row['CLP']/tot)) 

    #Recorre todo el diccionario resultante en busca de valores numéricos a formatear (poner puntos y símbolos de moneda con fmtoEntero())
    for clave, valor in nctxt.items():

        #Si el valor encontrado es una lista entonces es un valor a formatear. El primer elemenro es el número y el segundo su unidad
            #simpremente remplaza la lista por el falor ya formateado
            #Ej: [1000999000,4  ,  'USD'] -> 'US $1.001 millones'
        if type(valor) == list:
            nctxt[clave] = fmtoEntero(valor[0], valor[1])
    
    #Retorna el diccionario creado a partir del dataframe
    return nctxt



#Importa datos especificados por solicitante para cada región, retorna un diccionario de diccionarios regionales con los datos
def impAdicionalesReg ():
    addReg = pd.read_excel(io = 'datosAdicionales.xlsx') #requiere también openpyxl
    return addReg.set_index('region').to_dict('index')
