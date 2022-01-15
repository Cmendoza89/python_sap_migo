from tkinter import Tk, Label, Button, Frame,  messagebox, filedialog, ttk, Scrollbar, VERTICAL, HORIZONTAL
from tkinter import *
import pandas as pd
from openpyxl import Workbook, load_workbook
import shutil
from seaborn import load_dataset
import win32com.client
import sys
import subprocess
import time
import os



class  Carga : 
	global indica
	global tabla
	global ccdatos
	global cdatos
	global salir_login
	global login
	global usuario_info
	global clave_info
	

	

		
	def  __init__ ( self , connection , application, session) :
		 self .connection = connection
		 self .application = application
		 self .session = session
		 
		 


	ventana = Tk()
	ventana.config(bg='black')
	ventana.geometry('600x800')
	ventana.minsize(width=600, height=400)
	ventana.title('CARGA DE MEDIDORES _ CRISTHIAN MENDOZA')

	ventana.columnconfigure(0, weight = 25)
	ventana.rowconfigure(0, weight= 25)
	ventana.columnconfigure(0, weight = 1)
	ventana.rowconfigure(1, weight= 1)

	frame1 = Frame(ventana, bg='gray26')
	frame1.grid(column=0,row=0,sticky='nsew')
	frame2 = Frame(ventana, bg='gray26')
	frame2.grid(column=0,row=1,sticky='nsew')

	frame1.columnconfigure(0, weight = 1)
	frame1.rowconfigure(0, weight= 1)

	frame2.columnconfigure(0, weight = 1)
	frame2.rowconfigure(0, weight= 1)
	frame2.columnconfigure(1, weight = 1)
	frame2.rowconfigure(0, weight= 1)

	frame2.columnconfigure(2, weight = 1)
	frame2.rowconfigure(0, weight= 1)

	frame2.columnconfigure(3, weight = 2)
	frame2.rowconfigure(0, weight= 1)


	def abrir_archivo():

		global df
		global freq
		

		archivo = filedialog.askopenfilename(initialdir ='/', 
											title='Selecione archivo', 
											filetype=(('xls files', '*.xls*'),('All files', '*.*')))
		indica['text'] = archivo


		datos_obtenidos = indica['text']
		try:
			archivoexcel = r'{}'.format(datos_obtenidos)
			df = pd.read_excel(archivoexcel)
			

			#agregar nueva con datos de egreso e id de proyecto

			df['CODDOC']=df.apply(lambda x:'DOC:%s_ID:%s' % (x['CODDOC'],x['Unnamed: 17']),axis=1)
			
			#eliminar columnaS NO UTILES
			df= df.drop(['Unnamed: 17'], axis=1)
			df= df.drop(['ACTION'], axis=1)
			df= df.drop(['REFDOC'], axis=1)
			df= df.drop(['BUDAT'], axis=1)
			df= df.drop(['ERFMG'], axis=1)
			df= df.drop(['UMCHA'], axis=1)
			df= df.drop(['ERFME'], axis=1)


			#agregar nueva columna concatenada al df
			

			df['CONCATENACION']=df.apply(lambda x:'%s%s%s%s%s' % (x['BWART'],x['MAKTX'],x['UMNAME1'],x['UMLGOBE'],x['CODDOC']),axis=1)
			
			#ordenar datos
			df = df.sort_values('CONCATENACION')
			df.head()

			#obtener frecuencia de datos

			freq = df['CONCATENACION'].value_counts() 
			
			freq = freq.sort_values()
			freq.head()
			
			print(freq) 

			#indice de datos para carga
			#df['INDICE']=df.apply(lambda x:'%s_%s_%s_%s_%s' % (x['BWART'],x['MAKTX'],x['UMNAME1'],x['UMLGOBE'],x['CODDOC']),axis=1)
			#df.set_index('INDICE',inplace=True)
			#print (df)

			#dfw=df.drop_duplicates(subset=['INDICE'])
			#dfw.drop(['BWART', 'BLDAT', 'MAKTX'], axis=1)
			
		except ValueError:
			messagebox.showerror('Informacion', 'Formato incorrecto')
			return None

		except FileNotFoundError:
			messagebox.showerror('Informacion', 'El archivo esta \n erroneo')
			return None

		tabla.delete(*tabla.get_children())

		tabla['column'] = list(df.columns)
		tabla['show'] = "headings"  #encabezado
	     

		for columna in tabla['column']:
			tabla.heading(columna, text= columna)
		

		df_fila = df.to_numpy().tolist()
		cont = 0
		for fila in df_fila:
			tabla.insert('', 'end', values =fila)
			cont = cont + 1


		cdatos['text'] = cont


	def Limpiar():
		tabla.delete(*tabla.get_children())


	def carga():
		path =r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
		subprocess.Popen(path)
		time.sleep(5)

		SapGuiAuto = win32com.client.GetObject("SAPGUI")

		if not type (SapGuiAuto)== win32com.client.CDispatch:
			return

		application = SapGuiAuto.GetScriptingEngine
		#connection = application.OpenConnection ("05 UTQ CALIDAD ISU")
		connection = application.OpenConnection ("01 UTP PRODUCCION ISU")

		if not type (connection) == win32com.client.CDispatch:
			application = None
			SapGuiAuto = None
			return#

		session = connection.Children(0)
		if not type (session) == win32com.client.CDispatch:
			connection = None
			application = None
			SapGuiAuto = None
			return

		session.findById("wnd[0]").maximize
		session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario_info
		session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = clave_info
		session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
		session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 15
		session.findById("wnd[0]").sendVKey (0)
		session.findById("wnd[0]").sendVKey (0)
		session.findById("wnd[0]/tbar[0]/okcd").text = "migo"
		session.findById("wnd[0]").sendVKey (0)
		session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").topNode = "          1"
		

		#proceso para extraer datos del dataframe para carga

		global series
		global series2
		df_fila = freq.to_numpy().tolist()
		df_df = df['CONCATENACION'].drop_duplicates() 
		cont = 0
		for fila in df_fila:
			
			tabla.insert('', 'end', values =fila)
			ifila=df_df.iloc[cont] # el iloc controla la posición		
			df_mask=df['CONCATENACION']==ifila
			filtered_df = df[df_mask]
			filtered_df2= filtered_df.iloc[0]# el iloc controla la posición	
			
			print(cont)
			print(filtered_df)
			print(filtered_df2) #estos son los datos que se ingresan 
			#filtered_df['SERIALNO_01'].to_excel('D:\series.xlsx', header=None, index = False)# genera el documento excel para las series en SAP
			series= filtered_df['SERIALNO_01'].count()
			print(series)


			cont = cont + 1

			#traspaso 541 o 542 o 301 o 311[]
			session.findById("wnd[0]").maximize
			 
			
			#usamos el try para que haga el intento de cambiar desde entrada de mercancias hacia traspaso en el caso de no poder o encontrarse en traspaso significa que ya esta listo y salta al axcept
			try:

				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/cmbGODYNPRO-ACTION").key = "A08"
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").text = filtered_df2['BWART']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").caretPosition = 3
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT.").select
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-BWART").text = filtered_df2['BWART']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-BWART").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-BWART").caretPosition = 3
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS").select
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").caretPosition = 0
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").text = filtered_df2['MAKTX']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").caretPosition = 9
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-NAME1").text = filtered_df2['UMNAME1']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-LGOBE").text = filtered_df2['UMLGOBE']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-LGOBE").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-LGOBE").caretPosition = 4
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-CHARG").text = filtered_df2['CHARG']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-UMCHA").text = filtered_df2['CHARG']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-UMCHA").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-UMCHA").caretPosition = 5
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/subSUB_STOCK_IDENTIFIER_RIGHT:SAPLMIGO:0396/ctxtGOITEM-UMMAT_VENDORNAME").text = filtered_df2['UMMAT_VENDORNAME']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/subSUB_STOCK_IDENTIFIER_RIGHT:SAPLMIGO:0396/ctxtGOITEM-UMMAT_VENDORNAME").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/subSUB_STOCK_IDENTIFIER_RIGHT:SAPLMIGO:0396/ctxtGOITEM-UMMAT_VENDORNAME").caretPosition = 10
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/txtGODYNPRO-ERFMG").text = series
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/txtGODYNPRO-ERFMG").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/txtGODYNPRO-ERFMG").caretPosition = 1
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT.").select
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL").select
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-BKTXT").text = filtered_df2['CODDOC']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-SGTXT").text = filtered_df2['CODDOC']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-BKTXT").caretPosition = 3
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/btnOK_SER_REF").press
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL").select
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/btnOK_SER_REF").press
				session.findById("wnd[1]/usr/btnBUTTON_1").press
				session.findById("wnd[0]/usr/ctxtP_FILE").text = "D:\series.xlsx"
				session.findById("wnd[0]/usr/ctxtP_FILE").caretPosition = 14
				session.findById("wnd[0]/tbar[1]/btn[8]").press

				
			
			except:
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").text = filtered_df2['BWART']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_FIRSTLINE:SAPLMIGO:0010/ctxtGODEFAULT_TV-BWART").caretPosition = 3
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT.").select
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-BWART").text = filtered_df2['BWART']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-BWART").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/ctxtGOITEM-BWART").caretPosition = 3
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS").select
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").caretPosition = 0
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").text = filtered_df2['MAKTX']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-MAKTX").caretPosition = 9
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-NAME1").text = filtered_df2['UMNAME1']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-LGOBE").text = filtered_df2['UMLGOBE']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-LGOBE").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-LGOBE").caretPosition = 4
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-CHARG").text = filtered_df2['CHARG']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-UMCHA").text = filtered_df2['CHARG']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-UMCHA").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/ctxtGODYNPRO-UMCHA").caretPosition = 5
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/subSUB_STOCK_IDENTIFIER_RIGHT:SAPLMIGO:0396/ctxtGOITEM-UMMAT_VENDORNAME").text = filtered_df2['UMMAT_VENDORNAME']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/subSUB_STOCK_IDENTIFIER_RIGHT:SAPLMIGO:0396/ctxtGOITEM-UMMAT_VENDORNAME").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/subSUB_STOCK_IDENTIFIER_RIGHT:SAPLMIGO:0396/ctxtGOITEM-UMMAT_VENDORNAME").caretPosition = 10
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/txtGODYNPRO-ERFMG").text = series
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/txtGODYNPRO-ERFMG").setFocus
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_TRANS/ssubSUB_TS_GOITEM_TRANS:SAPLMIGO:0390/txtGODYNPRO-ERFMG").caretPosition = 1
				session.findById("wnd[0]").sendVKey (0)
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT.").select
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL").select
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-BKTXT").text = filtered_df2['CODDOC']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-SGTXT").text = filtered_df2['CODDOC']
				session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0112/txtGOHEAD-BKTXT").caretPosition = 3
				
				series2 = filtered_df['SERIALNO_01'].to_numpy().tolist()

				cont2 = -1

				

				print (series)

				for fila in range (series):
				

					seriesxxx = filtered_df['SERIALNO_01'].iloc[cont2]
					filtered_df['SERIALNO_01'].to_excel('D:\series.xlsx', header=None, index = False)

					#session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,1]").text = seriesxxx
					session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,0]").text = seriesxxx
					session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,0]").caretPosition = 6
					session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL").verticalScrollbar.position = cont2
					cont2 = cont2 + 1
					print (cont2)
					print (seriesxxx)



				#session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/tblSAPLMIGOTV_GOSERIAL/txtGOSERIAL-SERIALNO[0,1]").caretPosition = 4
				#session.findById("wnd[0]").sendVKey (0)


				#session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/btnOK_SER_REF").press
				#session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL").select
				#session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007/subSUB_ITEMDETAIL:SAPLMIGO:0303/subSUB_DETAIL:SAPLMIGO:0305/tabsTS_GOITEM/tabpOK_GOITEM_SERIAL/ssubSUB_TS_GOITEM_SERIAL:SAPLMIGO:0360/btnOK_SER_REF").press
				#session.findById("wnd[1]/usr/btnBUTTON_1").press
				#session.findById("wnd[0]/usr/ctxtP_FILE").text = "D:\series.xlsx"
				#session.findById("wnd[0]/usr/ctxtP_FILE").caretPosition = 14
				#session.findById("wnd[0]/tbar[1]/btn[8]").press




	

	def ventana_inicio():
	    global ventana_principal
	    global nombre_usuario
	    global clave
	    global entrada_clave
	    global entrada_nombre
	    nombre_usuario = StringVar() #DECLARAMOS "string" COMO TIPO DE DATO PARA "nombre_usuario"
	    clave = StringVar() #DECLARAMOS "sytring" COMO TIPO DE DATO PARA "clave"
	    pestas_color="DarkGrey"
	    ventana_principal=Tk()
	    ventana_principal.geometry("300x250")#DIMENSIONES DE LA VENTANA
	    ventana_principal.title("Login SAP")#TITULO DE LA VENTANA
	    etiqueta_nombre = Label(ventana_principal, text="Nombre de usuario * ")
	    etiqueta_nombre.pack()
	    entrada_nombre = Entry(ventana_principal, textvariable=nombre_usuario) #ESPACIO PARA INTRODUCIR EL NOMBRE.
	    entrada_nombre.pack()
	    etiqueta_clave = Label(ventana_principal, text="Contraseña * ")
	    etiqueta_clave.pack()
	    
	    #entrada_clave = ttk.Entry(ventana_principal)


	    entrada_clave = Entry(ventana_principal, textvariable=clave, show='ok') #ESPACIO PARA INTRODUCIR LA CONTRASEÑA.
	    entrada_clave.pack()
	    

	    Label(ventana_principal, text="").pack()
	    Button(ventana_principal, text="Aceptar", width=10, height=2, bg="LightGreen", command = login).pack() #BOTÓN "OK"
	    Button(ventana_principal, text="Salir", width=10, height=2, bg="LightGreen", command = salir_login).pack() #BOTÓN "OK"



	    ventana_principal.mainloop()

	def salir_login(): 


	    ventana_principal.destroy()	

	def login(): 
		global usuario_info
		global clave_info

		usuario_info = entrada_nombre.get()
		clave_info = entrada_clave.get()
		Label(ventana_principal, text="Usuario registrado", fg="green", font=("calibri", 11)).pack()
		
	    
	tabla = ttk.Treeview(frame1 , height=10)
	tabla.grid(column=0, row=0, sticky='nsew')

	ladox = Scrollbar(frame1, orient = HORIZONTAL, command= tabla.xview)
	ladox.grid(column=0, row = 1, sticky='ew') 

	ladoy = Scrollbar(frame1, orient =VERTICAL, command = tabla.yview)
	ladoy.grid(column = 1, row = 0, sticky='ns')

	tabla.configure(xscrollcommand = ladox.set, yscrollcommand = ladoy.set)

	estilo = ttk.Style(frame1)
	estilo.theme_use('clam') #  ('clam', 'alt', 'default', 'classic')
	estilo.configure(".",font= ('Arial', 14), foreground='red2')
	estilo.configure("Treeview", font= ('Helvetica', 12), foreground='black',  background='white')
	estilo.map('Treeview',background=[('selected', 'green2')], foreground=[('selected','black')] )


	boton1 = Button(frame2, text= 'Abrir', bg='Gray', command= abrir_archivo)
	boton1.grid(column = 0, row = 0, sticky='nsew', padx=10, pady=10)

	boton2 = Button(frame2, text= 'Carga', bg='Gray', command= carga)
	boton2.grid(column = 1, row = 0, sticky='nsew', padx=10, pady=10)

	boton3 = Button(frame2, text= 'Limpiar', bg='Gray', command= Limpiar)
	boton3.grid(column = 2, row = 0, sticky='nsew', padx=10, pady=10)

	boton4 = Button(frame2, text= 'Acceso SAP', bg='Gray', command= ventana_inicio)
	boton4.grid(column = 3, row = 0, sticky='nsew', padx=10, pady=10)



	cdatos = Label(frame2, fg= 'white', bg='gray26', text= '_', font= ('Arial',10,'bold') )
	cdatos.grid(column=5, row = 0)

	indica = Label(frame2, fg= 'white', bg='gray26', text= '_', font= ('Arial',10,'bold') )
	indica.grid(column=7, row = 0)

	ccdatos = Label(frame2, fg= 'white', bg='gray26', text= 'Transacciones:', font= ('Arial',10,'bold') )
	ccdatos.grid(column=4, row = 0)

	cindica = Label(frame2, fg= 'white', bg='gray26', text= 'Ruta:', font= ('Arial',10,'bold') )
	cindica.grid(column=6, row = 0)

	ventana.mainloop()