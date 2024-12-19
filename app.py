from tkinter import messagebox
import customtkinter as ctk
from tkinter import filedialog
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from separar_cia import separar_cias
import os

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Autenticación")
        self.root.geometry("400x300")

        ctk.CTkLabel(root, text="Email:").pack(pady=5)
        self.entry_email = ctk.CTkEntry(root)
        self.entry_email.pack(pady=5)

        ctk.CTkLabel(root, text="Contraseña:").pack(pady=5)
        self.entry_contraseña = ctk.CTkEntry(root, show="*")
        self.entry_contraseña.pack(pady=5)

        ctk.CTkButton(root, text="Iniciar sesión", command=self.verificar_credenciales).pack(pady=20)

    def verificar_credenciales(self):
        email = self.entry_email.get()
        contraseña = self.entry_contraseña.get()

        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(email, contraseña)
            server.quit()
            self.abrir_ventana_envio(email, contraseña)
        except smtplib.SMTPAuthenticationError:
            ctk.CTkLabel(self.root, text="Autenticación fallida. Inténtalo de nuevo.", fg_color="red").pack(pady=5)

    def abrir_ventana_envio(self, email, contraseña):
        self.root.destroy()

        self.nueva_ventana = ctk.CTk()
        self.nueva_ventana.title("Enviar correo")
        self.nueva_ventana.geometry("800x800")  # Aumenta el tamaño de la ventana si es necesario

        ctk.CTkLabel(self.nueva_ventana, text="Desde (Mail desde donde se enviará):").pack(pady=5)
        self.entry_from = ctk.CTkEntry(self.nueva_ventana, width=350)
        self.entry_from.pack(pady=5)

        ctk.CTkLabel(self.nueva_ventana, text="Asunto:").pack(pady=5)
        self.entry_asunto = ctk.CTkEntry(self.nueva_ventana, width=700)
        self.entry_asunto.pack(pady=5)

        ctk.CTkLabel(self.nueva_ventana, text="Cuerpo del mail:").pack(pady=5)
        self.textbox_cuerpo = ctk.CTkTextbox(self.nueva_ventana, width=700, height=150)
        self.textbox_cuerpo.pack(pady=10)

        # Casilla para incluir el nombre de la compañía en el asunto
        self.include_nombre_cia = ctk.CTkCheckBox(self.nueva_ventana, text="Incluir nombre de la compañía en el asunto")
        self.include_nombre_cia.pack(pady=5)

        # Botón para adjuntar archivo de destinatarios
        ctk.CTkButton(self.nueva_ventana, width=400, text="Adjuntar excel de destinatarios.", command=self.adjuntar_archivo_destinatarios).pack(pady=20)
        
        ctk.CTkButton(
            self.nueva_ventana,
            width=400,
            text="Adjuntar y procesar archivo a separar. (OPCIONAL)",
            command=self.separar_cias_individualmente,
            fg_color="gray",  # Color del texto y borde
            bg_color="lightblue"  # Color de fondo
        ).pack(pady=20)

        ctk.CTkButton(self.nueva_ventana, width=400, text="Seleccionar carpeta CIAS", command=self.carpeta_archivos_cias).pack(pady=20)

        # Botón para enviar correos
        ctk.CTkButton(self.nueva_ventana, width=400, text="ENVIAR", command=lambda: self.enviar_correo(email, contraseña)).pack(pady=30)

        self.nueva_ventana.mainloop()

    def adjuntar_archivo_destinatarios(self):
        self.archivo_destinatarios = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.archivo_destinatarios:
            self.data_frame_destinatarios = pd.read_excel(self.archivo_destinatarios)
            ctk.CTkLabel(self.nueva_ventana, text="Archivo de destinatarios cargado", fg_color="green").pack(pady=5)
        else:
            self.data_frame_destinatarios = None
            ctk.CTkLabel(self.nueva_ventana, text="No se seleccionó ningún archivo", fg_color="red").pack(pady=5)

    def separar_cias_individualmente(self):
        archivo_separacion = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("Zip files", "*.zip"), ("RAR files", "*.rar")])
        if archivo_separacion:
            separar_cias(self.nueva_ventana, archivo_separacion)
            ctk.CTkLabel(self.nueva_ventana, text="Archivo de separación procesado", fg_color="green").pack(pady=5)
        else:
            messagebox.showerror("Error", "No se ha seleccionado ningún archivo para separación.")

    def carpeta_archivos_cias(self):
        self.carpeta_cias = filedialog.askdirectory()
        if self.carpeta_cias:
            ctk.CTkLabel(self.nueva_ventana, text=f"Carpeta seleccionada: {self.carpeta_cias}", fg_color="green").pack(pady=5)
            print(self.carpeta_cias)
            archivos_en_carpeta = os.listdir(self.carpeta_cias)
            print(f"Archivos encontrados en la carpeta: {archivos_en_carpeta}")
        else:
            messagebox.showerror("Error", "No se ha seleccionado ninguna carpeta.")

    def enviar_correo(self, email, contraseña):
        asunto_base = self.entry_asunto.get()
        cuerpo = self.textbox_cuerpo.get("1.0", "end-1c")
        from_email = self.entry_from.get() or email  # Usa el correo especificado o el correo de autenticación

        # Verificar si el DataFrame está vacío
        if self.data_frame_destinatarios is None or self.data_frame_destinatarios.empty:
            messagebox.showerror("Error", "Debe adjuntar un archivo Excel válido.")
            return

        if not hasattr(self, 'carpeta_cias') or not self.carpeta_cias:
            messagebox.showerror("Error", "Debe seleccionar una carpeta de archivos de las compañías.")
            return

        try:
            smtp_server = 'smtp.gmail.com'
            smtp_port = 587
            smtp_username = email
            smtp_password = contraseña

            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(smtp_username, smtp_password)

            for index, row in self.data_frame_destinatarios.iterrows():
                correos = row.get('CORREOS')
                if isinstance(correos, str):
                    emails = correos.split(';')
                    codigo_compania = f"{int(row.get('CIA', '')):04}"
                    nombre_cia = row.get('NOMBRE', '')

                    # Personalizar el asunto
                    if self.include_nombre_cia.get():
                        asunto = f"{asunto_base} - {nombre_cia}"
                    else:
                        asunto = asunto_base

                    # Buscar todos los archivos correspondientes en la carpeta seleccionada
                    extensiones = ['xlsx', 'zip', 'rar']
                    for ext in extensiones:
                        archivo_adjunto = os.path.join(self.carpeta_cias, f'{codigo_compania}.{ext}')
                        if os.path.exists(archivo_adjunto):
                            msg = MIMEMultipart()
                            msg['From'] = from_email
                            msg['To'] = ', '.join(emails)
                            msg['Subject'] = asunto
                            msg.attach(MIMEText(cuerpo, 'plain'))

                            # Adjuntar el archivo
                            with open(archivo_adjunto, "rb") as adjunto:
                                part = MIMEApplication(adjunto.read(), Name=os.path.basename(archivo_adjunto))
                                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(archivo_adjunto)}"'
                                msg.attach(part)

                            server.sendmail(from_email, emails, msg.as_string())
                            mensaje = f"Correo enviado a: {', '.join(emails)} para compañía {codigo_compania} con archivo {os.path.basename(archivo_adjunto)}"
                            print(mensaje)
                            ctk.CTkLabel(self.nueva_ventana, text=mensaje, fg_color="green").pack(pady=5)

            server.quit()
            print("Proceso terminado.")
            ctk.CTkLabel(self.nueva_ventana, text="Proceso finalizado", fg_color="blue").pack(pady=10)
            messagebox.showinfo("Éxito", "Correos enviados exitosamente.")
        except Exception as e:
            print(f"Error al enviar correos: {e}")
            messagebox.showerror("Error", f"Error al enviar correos: {e}")

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    app = App(root)
    root.mainloop()
