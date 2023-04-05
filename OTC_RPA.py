# %%
import pandas as pd
import datetime
datecomun = str(datetime.datetime.today()).split()[0]
##
fecha = datetime.datetime.strptime(datecomun, "%Y-%m-%d").strftime("%d-%m-%Y")
date = str(datetime.datetime.today()).split()[0].replace("-","")
##
#import pyodbc
#
#cnxn_str = ("Driver={SQL Server};"
#            "Server=SRVBIROS02,1433;"
#            "Database=Acsa_Riesgo_DataWarehouse;"
#            "UID=usrReportingDatawarehouse;"
#            "PWD=8bWAY48c.2023;")
#cnxn = pyodbc.connect(cnxn_str)
#
#query = '''declare @fecha date
#select @fecha = MAX(Fecha)  from Compensacion.SimulacionContrato
#;With contratosOTC as (select *
#from Compensacion.SimulacionContrato  OTC
#	where otc.Fecha= @fecha and OTC.TipoContratoID in (23,24)
#),
#ultimos as (
#select ContratoID, fecha, MAX(timestamp) TS from Compensacion.MovimientoPrecio
#where Fecha= @fecha
#group by ContratoID, Fecha
#)
#select *
#--otc.ContratoID, otc.ContratoDescripcion, mp.PrecioAjuste
#from Compensacion.MovimientoPrecio mp
#inner join contratosOTC otc ON mp.Fecha = otc.fecha and otc.ContratoID = mp.ContratoID
#inner join ultimos u on u.Fecha = mp.Fecha and u.ContratoID = mp.ContratoID and u.TS = mp.TimeStamp
#order by mp.ContratoID'''
#
#df = pd.read_sql(query, cnxn)
#
#
#
#
## %%
#df["Futuro"] = df["ContratoDescripcion"].str.split().str[2]
#df = df.rename(columns={"ContratoDescripcion":"Posición OTC", "PrecioAjuste":"Precio"})
#otcdw = df[['Posición OTC','Futuro','Precio']]
#otcdw = otcdw[otcdw.Precio > 20]
#otcdw = df
filename = f"OTC {fecha}.xlsx"

# %%
#otcdw.to_excel(filename, index = False)
#token = "xoxb-25914778016-5021904796054-f24qrxBUbGfoa2Elu2mEkrXs"
#import slack
#from slack_sdk import WebClient 
#from slack_sdk.errors import SlackApiError 
#client = slack.WebClient(token=token)
#
#
#
#
#response = client.files_upload(
#    channels="#acyrsa-rpa-controles",
#    file=filename,
#    initial_comment= "Test"
#)
#
from selenium import webdriver
import pyautogui as pg
from datetime import datetime
import time
import datetime
import pyperclip as pc
    

now = datetime.datetime.now()


login=(1089,139)
entrar = (670,670)
selector_operacion = (370,186)
boton_ejecutar = (517,186)
boton_calculos = (117,210)
boton_guardar = (200,210)

editar=(1080,284)
cuadrado = (800,550)
config = (1031,480)
remove = (1250,508)
upload = (1170,552)
explorar = (257,50)

#operaciones=pg.locateCenterOnScreen("operaciones.png")
#ejecutar = pg.locateCenterOnScreen("ejecucion.png")
##selector_operacion = pg.locateCenterOnScreen("selecoper.png")
#boton_ejecutar = pg.locateCenterOnScreen("botonejecutar.png")
#boton_calculos = pg.locateCenterOnScreen("calculos.png")
#boton_guardar = pg.locateCenterOnScreen("guardar.png")

driver = webdriver.Chrome()
driver.get('https://www.jotform.com/')
driver.maximize_window()
time.sleep(10)
pg.click(login)
time.sleep(5)
pc.copy("mmartin@argentinaclearing.com.ar")
time.sleep(1)
pg.hotkey("ctrl","v")
time.sleep(2)
pg.press("tab")
time.sleep(1)
pc.copy("Marce-120195")
pg.hotkey("ctrl","v")
time.sleep(10)
pg.press("enter")
time.sleep(8)
pg.moveTo(editar)
time.sleep(4)
pg.click()
time.sleep(6)
pg.click(cuadrado)
time.sleep(6)
pg.click(config)
time.sleep(6)
pg.click(remove)
time.sleep(6)
pg.click(upload)
time.sleep(6)
pg.click(explorar)
time.sleep(2)
pg.click(explorar)
time.sleep(2)
pc.copy(r"C:\Users\aggonzalez\Downloads")
pg.hotkey("ctrl","v")
time.sleep(1)
pg.press("enter")
time.sleep(4)
pg.click(282,416)
        
        
pc.copy(filename)
pg.hotkey("ctrl","v")
time.sleep(4)
pg.press("enter")

publicar=(812,208)
time.sleep(6)
pg.click(publicar)





time.sleep(19)


