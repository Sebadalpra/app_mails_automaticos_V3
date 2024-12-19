import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd

# Definir la información del servidor SMTP y las credenciales
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = ''
smtp_password = ''  # Contraseña de aplicación

# Cargar un archivo de Excel
tabla_destinatarios = pd.read_excel('mails_cias.xlsx', sheet_name='Sheet1')

# Iterar a través de la tabla de destinatarios y enviar correos individuales
for index, row in tabla_destinatarios.iterrows():
    emails = row['CORREOS'].split(';')  # Dividir las direcciones de correo por punto y coma
    codigo_compania = row['CIA']
    nombre_compania = row['NOMBRE']
    
    for correo in emails:
        # Crear un objeto MIMEMultipart para el correo
        msg = MIMEMultipart()
        msg['From'] = 'jym@ssn.gob.ar'
        msg['To'] = correo.strip()  # Eliminar espacios en blanco alrededor de la dirección de correo
        msg['Subject'] = 'AGENDAR VENCIMIENTO Juicios y Mediaciones 2024-2 RGAA 39.6.4'

        # Mensaje que deseas enviar
        mensaje = f"""

        """

        # Agregar el mensaje al cuerpo del correo con codificación UTF-8
        msg.attach(MIMEText(mensaje, 'plain', 'utf-8'))

        # Conectar y autenticar con el servidor de correo
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)

        # Enviar el correo al destinatario actual
        server.sendmail(smtp_username, correo.strip(), msg.as_string())

        # Cerrar la conexión con el servidor de correo
        server.quit()

        print(f"Correo enviado a {correo.strip()} para {codigo_compania}")

print("Ejecución finalizada.")
