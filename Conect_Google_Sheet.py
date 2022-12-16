from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait

import gspread
from oauth2client.service_account import ServiceAccountCredentials

import pandas as pd

import time


path = "/chromedriver.exe"
Service = Service(executable_path=path)
driver = webdriver.Chrome(service=Service)
# driver.maximize_window()
driver.get("http://jvelazquez:Nacho123-@crm.telecentro.local/MembersLogin.aspx")
time.sleep(1)

driver.find_element(
    by="xpath", value='//*[@id="txtPassword"]').send_keys("Nacho123-")

driver.find_element(
    by="xpath", value='//*[@id="btnAceptar"]').send_keys(Keys.RETURN)

time.sleep(1)


driver.get(
    "http://crm.telecentro.local/Edificio/Gt_Edificio/BandejaEntradaDeRelevamiento.aspx?TituloPantalla=Descarga%20De%20Relevamiento&EstadoGestionId=5&TipoGestionId=3&TipoGestion=OPERACIONES%20DE%20RED%20-%20CIERRE%20DE%20RELEVAMIENTO")

time.sleep(1)

driver.find_element(
    by="xpath", value='//*[@id="btnBuscar"]').send_keys(Keys.RETURN)

time.sleep(3)

filas = len(driver.find_elements(
    by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr'))

columnas = len(driver.find_elements(
    by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr[1]/td'))

print(filas)
print(columnas)

mensaje = driver.find_element(
    by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table/tbody/tr[51]/td[14]/a')
print(mensaje.text)
print(mensaje.get_attribute('onmouseover'))

# ------------------------------------------------------

scope = ['https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive"]

credenciales = ServiceAccountCredentials.from_json_keyfile_name(
    "carga-de-bandeja-de-entrada-01ec277da545.json", scope)

cliente = gspread.authorize(credenciales)
# ------------- Crea y comparte la Google Sheet  -------------------------
#libro = cliente.create("AutocargaGestiones")
#libro.share("ignaciogproce3@gmail.com", perm_type="user", role="writer")
# ------------------------------------------------------------------------

hoja = cliente.open("AutocargaGestiones").sheet1

#hoja.update_cell(1, 2, 'Bingo!')
#hoja.update("A4", Titulo)


datos = []
for x in range(1, filas+1):
    for y in range(1, columnas+1):
        dato = driver.find_element(
            by="xpath", value='//*[@id="ctl00_ContentBody_grGestionTecEdificio"]/div/table/tbody/tr['+str(x)+"]/td["+str(y)+"]").text
        #print(dato, end="       ")
        if y == 7:
            direc = driver.find_element(
                by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table/tbody/tr['+str(x)+']/td[7]/a')
            altura = direc.get_attribute('onmouseover')[13:]
            comilla = altura.index("'")
            altura = altura[:comilla]
            datos.append(altura)
        else:
            datos.append(dato)
        if y == 14:
            obs = driver.find_element(
                by="xpath", value='/html/body/form/div[4]/div[4]/table[1]/tbody/tr[10]/td/table/tbody/tr/td/div/div/div/table/tbody/tr['+str(x)+']/td[14]/a')
            cadena = obs.get_attribute('onmouseover')[13:]
            comilla = cadena.index("'")
            cadena = cadena[:comilla]
            datos.append(cadena)


# print(datos)
serie = pd.Series(datos)
df = pd.DataFrame(serie.values.reshape(filas, columnas+1))
df.columns = ["N", "Gestion", "ID", "Nodo", "Zona", "Prioridad", "Direccion", "Subtipo", "Ult Visita",
              "Estado Edificio", "Cant Gestiones", "Usuario", "Contratista", "Bandeja Previa", "Observacion", "Nodo Gpon"]
print(df)

hoja.update([df.columns.values.tolist()] + df.values.tolist())

# -------------------------------------------------------


print("Inicio")

print("Fin")

input("Esperando que no se cierre webdriver: ")


"""


"""
