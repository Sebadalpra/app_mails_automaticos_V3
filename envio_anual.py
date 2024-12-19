import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

# Definir la información del servidor SMTP y las credenciales
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = 'mferro@ssn.gob.ar' 
smtp_password = ''  # Esto es una contraseña de aplicación que genera a partir de la verificación en 2 pasos.

# Cargar un archivo de Excel
tabla_destinatarios = pd.read_excel('contactos_infopro.xlsx', sheet_name='Sheet1')

# Iterar a través de la tabla de destinatarios y enviar correos individuales
for index, row in tabla_destinatarios.iterrows():
    email = row['CORREOS'].split(';')  # Dividir las direcciones de correo por punto y coma
    codigo_compania = row['CIA']
    nombre_compañia = row['NOMBRE']
    
    for correo in email:
        # Crear un objeto MIMEMultipart para el correo
        msg = MIMEMultipart()
        msg['From'] = smtp_username
        msg['To'] = correo.strip()  # Eliminar espacios en blanco alrededor de la dirección de correo
        msg['Subject'] = f'INFOPRO 2024'

        # Mensaje que deseas enviar
        mensaje = f"""
        <html>
        <body style="text-align: justify;">
        <p>Estimados;</p>

        <p>Por medio del presente le hacemos saber que ya está habilitado el sistema de <b>INFOPRO</b>
        para realizar la carga de la información de los distintos canales de Venta - <b>INFOPRO</b>
        (Período: 2023-2024) Comunicación SSN 3733 con fecha 29/08/2013.</p>

        <p>Cabe aclarar que como condición fundamental:</p>

        <ul>
            <li>Las Primas Emitidas Netas de Anulaciones de los <b>ramos Patrimoniales</b> informada en el
            sistema de <b>INFOPRO</b> debe coincidir con la informada al sistema <b>SINENSUP</b> al cierre del
            ejercicio a valores ajustados.</li>
            
            <li>Primas Emitidas netas de anulaciones de los <b>ramos de seguros de Personas</b>
            informada en el sistema de <b>INFOPRO</b> debe coincidir con la informada al sistema
            <b>SINENSUP</b> al cierre del ejercicio a valores ajustados.</li>
            
            <li>Primas Emitidas Netas de Anulaciones de los <b>ramos Patrimoniales</b> informada en el
            sistema de <b>INFOPRO</b> debe coincidir con la informada al sistema <b>DISTRIBUCION
            GEOGRAFICA</b> al cierre del ejercicio.</li>
            
            <li>Primas Emitidas Netas de Anulaciones de <b>los ramos de Seguros de Personas</b>
            informada en el sistema de <b>INFOPRO</b> debe coincidir con la informada al sistema
            <b>DISTRIBUCION GEOGRAFICA</b> al cierre del ejercicio.</li>
        </ul>

        <p>En la pantalla de carga están los link de todas las validaciones contempladas en el sistema y
        los códigos de las distintas jurisdicciones.</p>

        <p>La fecha de vencimiento de esta información será a los 45 días posteriores al vencimiento
        de la presentación de los estados contables (SINENSUP).</p>

        <p>Ante cualquier duda comunicarse con mferro@ssn.gob.ar ; vsabell@ssn.gob.ar</p>

        <p>Maria Florencia Ferro <br>
        Gerencia de Estudios y Estadísticas - Maria Florencia Ferro <br>
        mferro@ssn.gob.ar <br>
        Tel: +54 11 4338-4000 | Int 1559 <br>
        <a href="http://www.argentina.gob.ar/ssn">www.argentina.gob.ar/ssn</a> <br>
        Av. Julio A. Roca 721 | (C1067ABC) | CABA | República Argentina</p>
        </body>
        </html>
        """

        # Agregar el mensaje al cuerpo del correo
        msg.attach(MIMEText(mensaje, 'html'))

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
