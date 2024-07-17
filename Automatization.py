import xlwings as xw
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import re
import datetime

# Función para enviar correo
def enviar_correo(destinatario, asunto, mensaje):
    remitente = ''        #Colocar gmail desde donde se envia el e-mail
    password = ''         #Contraseña de aplicación
    msg = MIMEMultipart() #Creo variable tipo MIME
    msg['From'] = remitente
    msg['To'] = destinatario
    msg['Subject'] = asunto
    msg.attach(MIMEText(mensaje, 'plain'))
    server = smtplib.SMTP('smtp.gmail.com', 587) #Accedo al server de GMAIL
    server.starttls()
    server.login(remitente, password)
    text = msg.as_string()
    server.sendmail(remitente, destinatario, text)
    server.quit()
#Funcion para Validar Mail
def es_email_valido(email):
    patron = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    return re.match(patron, email) is not None

# Función para llenar el formulario web
def llenar_formulario(driver, datos):
    proceso = driver.find_element(By.ID, 'process')
    riesgo = driver.find_element(By.ID, 'tipo_riesgo')
    severidad = driver.find_element(By.ID, 'severidad')
    responsable = driver.find_element(By.ID, 'res')
    fecha = driver.find_element(By.ID, 'date')
    observacion = driver.find_element(By.ID, 'obs')
    proceso.send_keys(datos['proceso'])
    riesgo.send_keys(datos['riesgo'])
    severidad.send_keys(datos['severidad'])
    responsable.send_keys(datos['responsable'])
    fecha.send_keys(datos['fecha'])
    observacion.send_keys(datos['observacion'])
    submit_button = driver.find_element(By.ID, 'submit')  
    submit_button.click()

# Abre el archivo Excel
wb = xw.Book('Base Seguimiento Observ Auditoría al_30042021 - copia.xlsx')
sht = wb.sheets['Hoja1']
estados = sht.range('J2:J' + str(sht.cells.last_cell.row)).value #Recorro el archivo en la fila J y lo guardo

# Inicializa el WebDriver de Chrome
service = Service(r"chromedriver-win64\chromedriver.exe")
driver = webdriver.Chrome(service=service)
driver.get('https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG')
time.sleep(3) # espera en segundos para darle tiempo a Chrome de abrirse

# Procesa cada fila según el estado del proceso
for i, estado in enumerate(estados, start=2):
    if estado == "Regularizado":
        datos = {
            'proceso': sht.range(f'A{i}').value,  
            'riesgo': sht.range(f'C{i}').value,
            'severidad': sht.range(f'D{i}').value,
            'responsable': sht.range(f'G{i}').value,
            'fecha': sht.range(f'F{i}').value,
            'observacion': sht.range(f'B{i}').value,                     
        }
        for key, value in datos.items():
            if isinstance(value, datetime.datetime):
                datos[key] = value.strftime('%d-%m-%Y %H:%M:%S')
            elif isinstance(value, datetime.date):
                datos[key] = value.strftime('%d-%m-%Y')
        llenar_formulario(driver, datos)
    elif estado == "Atrasado":
        destinatario = sht.range(f' I{i}').value.strip()
        if es_email_valido(destinatario): #Corroboro que el mail sea correcto
            asunto = "Proceso Atrasado"
            mensaje = f"El proceso {sht.range(f'A{i}').value} está atrasado. Observación: {sht.range(f'B{i}').value}, Fecha de compromiso: {sht.range(f'F{i}').value}."
            enviar_correo(destinatario, asunto, mensaje)
        else:
            print(f"Dirección de correo no válida: {destinatario}")

#Cierra el archivo Excel y el navegador
wb.close()
driver.quit()
