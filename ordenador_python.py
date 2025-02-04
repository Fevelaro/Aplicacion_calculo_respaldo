import pandas as pd
#import xlrd
import os
#import matplotlib.pyplot as plt
import numpy as np
import xlsxwriter
from datetime import datetime
from vincent.colors import brews
np.seterr(divide='ignore', invalid='ignore')
#######################################################################################################################################################
## Definición de funciones
def extraer_datos(ruta):	
	try:
		df=pd.read_excel(ruta, sheet_name='Detalle remuneración', header=0) ## El parámetro sheet_name identifica el nombre de la hoja, de no coincidir no extrae datos de allí
	except ValueError:
		input('\nEl archivo no tiene una hoja con un nombre conocido (Resumen por Empresa)\n')

	fecha=pd.to_datetime(df["Fecha"],format='%d/%m/%Y')
	df['Mes']=fecha.dt.month.values
	df['Año']=fecha.dt.year.values
	df.drop(['Fecha'], axis=1,inplace=True)
	df.drop(['Hora'], axis=1,inplace=True)
	return df

def escribidor(hojaf1,hojaf2,hojaf3,anem,antec,anug,grafh1,grafh2,grafh3,ruta): ### Escribe los dataframe en un libro excel.
	nombre=ruta+'Resumen_Anual.xlsx'
	writer = pd.ExcelWriter(nombre, engine='xlsxwriter')
	grafh1.to_excel(writer, sheet_name='Resumen_mes_empresa') ##nombra las hojas del libro
	grafh2.to_excel(writer, sheet_name='Resumen_mes_tecnología') ##nombra las hojas del libro
	grafh3.to_excel(writer, sheet_name='Resumen_mes_por_unidad') ##nombra las hojas del libro

	anem.to_excel(writer, sheet_name='Anual Tecnología') ##nombra las hojas del libro
	antec.to_excel(writer, sheet_name='Anual Empresa') ##nombra las hojas del libro
	anug.to_excel(writer, sheet_name='Anual Unidad Generadora') ##nombra las hojas del libro
#	hojaf1.to_excel(writer, sheet_name='Empresa') ##nombra las hojas del libro
#	hojaf2.to_excel(writer, sheet_name='Tecnología') ##nombra las hojas del libro
#	hojaf3.to_excel(writer, sheet_name='Unidad Generadora') ##nombra las hojas del libro

	wb  = writer.book
	formato_remu = wb.add_format({'num_format': '[$$-409]#,##0.00'}) ##Define los formatos a usar según tipo de columna
	formato_per = wb.add_format({'num_format': '0.00%'})

	ws = writer.sheets['Resumen_mes_empresa'] ##Aplica los formato de columna según corresponda
	##########Aplica formato a hoja 1
	ws.set_column(0,0,20, None)
	filas,columnas=grafh1.shape
	for col in range(columnas):
		ws.set_column(col+1, col+1, 10, formato_per)

	##########Aplica formato a hoja 2
	ws = writer.sheets['Resumen_mes_tecnología'] ##Aplica los formato de columna según corresponda
	ws.set_column(0,0,30, None)
	filas,columnas=grafh2.shape
	for col in range(columnas):
		ws.set_column(col+1, col+1, 10, formato_per)

	##########Aplica formato a hoja 3
	ws = writer.sheets['Resumen_mes_por_unidad'] ##Aplica los formato de columna según corresponda
	ws.set_column(0,0,30, None)
	filas,columnas=grafh3.shape
	for col in range(columnas):
		ws.set_column(col+1, col+1, 10, formato_per)


	############Aplica formato a hoja4
	ws = writer.sheets['Anual Tecnología'] ##Aplica los formato de columna según corresponda
	ws.set_column(0,0,30, None)
	ws.set_column(1, 1, 18, formato_remu)
	ws.set_column(2, 2, 18, formato_remu)
	ws.set_column(3, 3, 10, formato_per)

	###Inserta gráfico en hoja 4
	tec_graficadas=len(anem)
	ws = writer.sheets['Anual Tecnología']
	chart=wb.add_chart({'type':'column'})
	chart.add_series({'categories':['Anual Tecnología',1,0,tec_graficadas,0],
		'values':['Anual Tecnología',1,3,tec_graficadas,3],
		'name':'Porcentaje de remuneración alcanzada en el año',})
	chart.set_legend({'none': True})
	ws.insert_chart('F1',chart)

	############Aplica formato a hoja5
	ws = writer.sheets['Anual Empresa'] ##Aplica los formato de columna según corresponda
	ws.set_column(0,0,25, None)
	ws.set_column(1, 1, 18, formato_remu)
	ws.set_column(2, 2, 18, formato_remu)
	ws.set_column(3, 3, 10, formato_per)

	###Inserta gráfico en hoja 5
	em_graficadas=len(antec)
	ws = writer.sheets['Anual Empresa']
	chart=wb.add_chart({'type':'column'})
	chart.add_series({'categories':['Anual Empresa',1,0,em_graficadas,0],
		'values':['Anual Empresa',1,3,em_graficadas,3],
		'name':'Porcentaje de remuneración alcanzada en el año',})
	chart.set_legend({'none': True})
	ws.insert_chart('F1',chart)

	############Aplica formato a hoja6
	ws = writer.sheets['Anual Unidad Generadora'] ##Aplica los formato de columna según corresponda
	ws.set_column(0,0,25, None)
	ws.set_column(1, 1, 18, formato_remu)
	ws.set_column(2, 2, 18, formato_remu)
	ws.set_column(3, 3, 10, formato_per)

	###Inserta gráfico en hoja 6
	ug_graficadas=len(anug)
	ws = writer.sheets['Anual Unidad Generadora']
	chart=wb.add_chart({'type':'column'})
	chart.add_series({'categories':['Anual Unidad Generadora',1,0,ug_graficadas,0],
		'values':['Anual Unidad Generadora',1,3,ug_graficadas,3],
		'name':'Porcentaje de remuneración alcanzada en el año',})
	chart.set_legend({'none': True})
	ws.insert_chart('F1',chart)

#	Inserción de graficos en hojas 1,2 y3

	#grafh1.to_excel(writer, sheet_name='Resumen_mes_empresa')
	filas,columnas=grafh1.shape#filas, columnas (6 y 12)
	ws = writer.sheets['Resumen_mes_empresa']
	chart=wb.add_chart({'type':'line'})
	for fila in range(filas):
	    chart.add_series({
	        'name':       ['Resumen_mes_empresa', fila+3 , 0],
	        'categories': ['Resumen_mes_empresa', 1, 1, 1, columnas],
	        'values':     ['Resumen_mes_empresa', fila+3, 1, fila+3, columnas],
#	        'fill':       {'color': brews['Spectral'][fila]},
	        'gap':        300,
	    })

	chart.set_x_axis({'name': 'Año-Mes'})
	chart.set_y_axis({'name': 'Remuneración [%]', 'major_gridlines': {'visible': False}})
	chart.set_size({'x_scale': 2, 'y_scale': 1})
	ws.insert_chart('A'+str(filas+4),chart)


	filas,columnas=grafh2.shape#filas, columnas (6 y 12)
	ws = writer.sheets['Resumen_mes_tecnología']
	chart=wb.add_chart({'type':'line'})
	for fila in range(filas):
	    chart.add_series({
	        'name':       ['Resumen_mes_tecnología', fila+3 , 0],
	        'categories': ['Resumen_mes_tecnología', 1, 1, 1, columnas],
	        'values':     ['Resumen_mes_tecnología', fila+3, 1, fila+3, columnas],
#	        'fill':       {'color': brews['Spectral'][fila]},
	        'gap':        300,
	    })

	chart.set_x_axis({'name': 'Año-Mes'})
	chart.set_y_axis({'name': 'Remuneración [%]', 'major_gridlines': {'visible': False}})
	chart.set_size({'x_scale': 2, 'y_scale': 1})
	ws.insert_chart('A'+str(filas+4),chart)

	filas,columnas=grafh3.shape#filas, columnas (6 y 12)
	ws = writer.sheets['Resumen_mes_por_unidad']
	chart=wb.add_chart({'type':'line'})
	for fila in range(filas):
	    chart.add_series({
	        'name':       ['Resumen_mes_por_unidad', fila+3 , 0],
	        'categories': ['Resumen_mes_por_unidad', 1, 1, 1, columnas],
	        'values':     ['Resumen_mes_por_unidad', fila+3, 1, fila+3, columnas],
#	        'fill':       {'color': brews['Spectral'][fila]},
	        'gap':        300,
	    })

	chart.set_x_axis({'name': 'Año-Mes'})
	chart.set_y_axis({'name': 'Remuneración [%]', 'major_gridlines': {'visible': False}})
	chart.set_size({'x_scale': 2, 'y_scale': 1})
	ws.insert_chart('A'+str(filas+4),chart)

	writer.save()


def h1(df):
	mes=int(df['Mes'].mean())
	año=int(df['Año'].mean())
	df_gbE=df.groupby(['Empresa']).agg(sum)
	num=np.array(df_gbE['Remuneración Real'].values,dtype='float')
	den=np.array(df_gbE['Remuneración Ideal'].values,dtype='float')
	porcentaje=num/den
	df_gbE['Porcentaje']=porcentaje
	df_gbE['Mes']=mes
	df_gbE['Año']=año
	df_gbE.reset_index(inplace=True)
	return df_gbE

def h2(df):
	mes=int(df['Mes'].mean())
	año=int(df['Año'].mean())
	df_gbT=df.groupby(['Tecnología']).agg(sum)
	num=np.array(df_gbT['Remuneración Real'].values,dtype='float')
	den=np.array(df_gbT['Remuneración Ideal'].values,dtype='float')
	porcentaje=num/den
	df_gbT['Porcentaje']=porcentaje
	df_gbT['Mes']=mes
	df_gbT['Año']=año
	df_gbT.reset_index(inplace=True)
#	input(df_gbT)
	return df_gbT

def h3(df):
	mes=int(df['Mes'].mean())
	año=int(df['Año'].mean())
	df_gbUG = df.drop(df[df['Empresa']!='COLBUN'].index)
	df_gbUG=df_gbUG.groupby(['Unidad Generadora']).agg(sum)
	num=np.array(df_gbUG['Remuneración Real'].values,dtype='float')
	den=np.array(df_gbUG['Remuneración Ideal'].values,dtype='float')
	porcentaje=num/den
	df_gbUG['Porcentaje']=porcentaje
	df_gbUG['Mes']=mes
	df_gbUG['Año']=año
	df_gbUG.reset_index(inplace=True)
	return df_gbUG



############################################################################################################################################################################
###Comienza el programa

#Conociendo directorio del script
absFilePath = os.path.abspath(__file__)
ruta, filename = os.path.split(absFilePath)
ruta=ruta+'\\nuevos\\'
files = os.listdir(ruta) #Lista con archivos

##Llamando a la función que extrae los datos.
hojaf1=pd.DataFrame(columns=['Empresa','Remuneración Ideal','Remuneración Real','Porcentaje','Mes','Año'])
hojaf2=pd.DataFrame(columns=['Tecnología','Remuneración Ideal','Remuneración Real','Porcentaje','Mes','Año'])
hojaf3=pd.DataFrame(columns=['Unidad Generadora','Remuneración Ideal','Remuneración Real','Porcentaje','Mes','Año'])
for i in range(len(files)):
	rutaA=ruta+files[i]##Es la ruta de cada archivo
	name_arch=rutaA.split("\\")
	narch=name_arch.pop()
	df=extraer_datos(rutaA)
	hoja1=h1(df)
	hojaf1=pd.concat([hojaf1,hoja1],axis=0)
	hoja2=h2(df)
	hojaf2=pd.concat([hojaf2,hoja2],axis=0)
	hoja3=h3(df)
	hojaf3=pd.concat([hojaf3,hoja3],axis=0)

hojaf1=hojaf1.groupby(['Año','Mes','Empresa'], as_index=False).agg(max)
hojaf2=hojaf2.groupby(['Año','Mes','Tecnología'], as_index=False).agg(max)
hojaf3=hojaf3.groupby(['Año','Mes','Unidad Generadora'], as_index=False).agg(max)

anem=hojaf2.groupby(['Tecnología']).agg(sum)
num=np.array(anem['Remuneración Real'].values,dtype='float')
den=np.array(anem['Remuneración Ideal'].values,dtype='float')
porcentaje=num/den
anem['Porcentaje']=porcentaje
anem.drop(['Año','Mes'],axis=1,inplace=True)

antec=hojaf1.groupby(['Empresa']).agg(sum)
num=np.array(antec['Remuneración Real'].values,dtype='float')
den=np.array(antec['Remuneración Ideal'].values,dtype='float')
porcentaje=num/den
antec['Porcentaje']=porcentaje
antec.drop(['Año','Mes'],axis=1,inplace=True)

anug=hojaf3.groupby(['Unidad Generadora']).agg(sum)
num=np.array(anug['Remuneración Real'].values,dtype='float')
den=np.array(anug['Remuneración Ideal'].values,dtype='float')
porcentaje=num/den
anug['Porcentaje']=porcentaje
anug.drop(['Año','Mes'],axis=1,inplace=True)

hojaf1.set_index(['Empresa'],inplace=True)
hojaf2.set_index(['Tecnología'],inplace=True)
hojaf3.set_index(['Unidad Generadora'],inplace=True)

input(hojaf1)

grafh1=hojaf1.copy()
grafh1['Mes']=grafh1['Mes'].astype(str)
grafh1['Mes'] = grafh1['Mes'].replace({"1": '01', "2": '02',"3":'03',"4":'04',"5":'05',"6":'06',"7":'07',"7":'07',"8":'08',"9":'09'})
grafh1['Mesaño']=grafh1['Año'].astype(str)+'-'+grafh1['Mes'].astype(str)
grafh1.reset_index(inplace=True)
grafh1=grafh1.pivot(index='Empresa',columns='Mesaño')
grafh1.drop(['Remuneración Ideal','Remuneración Real','Mes','Año'],axis=1,inplace=True)
grafh1=grafh1.replace(np.nan,0)

grafh2=hojaf2.copy()
grafh2['Mes']=grafh2['Mes'].astype(str)
grafh2['Mes'] = grafh2['Mes'].replace({"1": '01', "2": '02',"3":'03',"4":'04',"5":'05',"6":'06',"7":'07',"7":'07',"8":'08',"9":'09'})
grafh2['Mesaño']=grafh2['Año'].astype(str)+'-'+grafh2['Mes'].astype(str)
grafh2.reset_index(inplace=True)
grafh2=grafh2.pivot(index='Tecnología',columns='Mesaño')
grafh2.drop(['Remuneración Ideal','Remuneración Real','Mes','Año'],axis=1,inplace=True)
grafh2=grafh2.replace(np.nan,0)

grafh3=hojaf3.copy()
grafh3['Mes']=grafh3['Mes'].astype(str)
grafh3['Mes'] = grafh3['Mes'].replace({"1": '01', "2": '02',"3":'03',"4":'04',"5":'05',"6":'06',"7":'07',"7":'07',"8":'08',"9":'09'})
grafh3['Mesaño']=grafh3['Año'].astype(str)+'-'+grafh3['Mes'].astype(str)
grafh3.reset_index(inplace=True)
grafh3=grafh3.pivot(index='Unidad Generadora',columns='Mesaño')
grafh3.drop(['Remuneración Ideal','Remuneración Real','Mes','Año'],axis=1,inplace=True)
grafh3=grafh3.replace(np.nan,0)

escribidor(hojaf1,hojaf2,hojaf3,anem,antec,anug,grafh1,grafh2,grafh3,ruta)
