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
    remitente = 'pentitodante2@gmail.com'
    password = 'owyj jodq lhnh xmgy'
    msg = MIMEMultipart()
    msg['From'] = remitente
    msg['To'] = destinatario
    msg['Subject'] = asunto
    msg.attach(MIMEText(mensaje, 'plain'))
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(remitente, password)
    text = msg.as_string()
    server.sendmail(remitente, destinatario, text)
    server.quit()
#Valido Mail
def es_email_valido(email):
    patron = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    return re.match(patron, email) is not None

# Función para llenar el formulario web
def llenar_formulario(driver, datos):
    campo1 = driver.find_element(By.ID, 'process')
    campo2 = driver.find_element(By.ID, 'tipo_riesgo')
    campo3 = driver.find_element(By.ID, 'severidad')
    campo4 = driver.find_element(By.ID, 'res')
    campo5 = driver.find_element(By.ID, 'date')
    campo6 = driver.find_element(By.ID, 'obs')
    campo1.send_keys(datos['campo1'])
    campo2.send_keys(datos['campo2'])
    campo3.send_keys(datos['campo3'])
    campo4.send_keys(datos['campo4'])
    campo5.send_keys(datos['campo5'])
    campo6.send_keys(datos['campo6'])
    campo6.send_keys(Keys.RETURN)
    submit_button = driver.find_element(By.ID, 'submit')  
    submit_button.click()
print("Formulario web definido")

# Abre el archivo Excel
wb = xw.Book('Archivo 1.xlsx')
sht = wb.sheets['Hoja1']
estados = sht.range('J2:J' + str(sht.cells.last_cell.row)).value
print("archivo excel abierto")

# Inicializa el WebDriver de Chrome
service = Service(r"chromedriver-win64\chromedriver.exe")
driver = webdriver.Chrome(service=service)
driver.get('https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG')
time.sleep(3) # espera en segundos para darle tiempo a Chrome de abrirse

# Procesa cada fila según el estado del proceso
for i, estado in enumerate(estados, start=2):
    if estado == "Regularizado":
        datos = {
            'campo1': sht.range(f'A{i}').value,  
            'campo2': sht.range(f'C{i}').value,
            'campo3': sht.range(f'D{i}').value,
            'campo4': sht.range(f'G{i}').value,
            'campo5': sht.range(f'F{i}').value,
            'campo6': sht.range(f'B{i}').value,                     
        }
        for key, value in datos.items():
            if isinstance(value, datetime.datetime):
                datos[key] = value.strftime('%Y-%m-%d %H:%M:%S')
            elif isinstance(value, datetime.date):
                datos[key] = value.strftime('%Y-%m-%d')
        llenar_formulario(driver, datos)
    elif estado == "Atrasado":
        destinatario = sht.range(f' I{i}').value.strip()
        if es_email_valido(destinatario):
            asunto = "Proceso Atrasado"
            mensaje = f"El proceso {sht.range(f'A{i}').value} está atrasado. Observación: {sht.range(f'B{i}').value}, Fecha de compromiso: {sht.range(f'F{i}').value}."
            enviar_correo(destinatario, asunto, mensaje)
        else:
            print(f"Dirección de correo no válida: {destinatario}")

# Cierra el archivo Excel y el navegador
wb.close()
driver.quit()