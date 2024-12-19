# APP enviar mails automaticos V3

## Esta app manda un archivo Excel, rar o zip a su correspondiente contacto.

Necesario:
-	Crear una contraseña a partir de la verificación en 2 pasos de google.
Para la creación de esta contraseña debemos:
1.	Dirigirnos a https://gmail.com/ con la sesión iniciada de la cuenta que utilizaremos para enviar los mails.
2.	Seleccionar nuestra foto de perfil que aparece en Gmail y luego: “Administrar tu Cuenta de Google”
3.	Donde dice “Buscar en la Cuenta de Google” escribimos “Verificación en 2 pasos” y seleccionamos la opción.
4.	Una vez allí, bajamos hasta la sección “Contraseñas de aplicaciones” y hacemos click en esa flecha.
5.	En “Nombre de la app” le ponemos un nombre como por ejemplo: “Automatización Mails”.
6.	Copiamos la contraseña y nos la guardamos en algún Word, ya que es la que utilizaremos siempre.

Proceso:
1.  Al abrir la aplicación, verás una pantalla de inicio de sesión donde deberás introducir tu correo electrónico y la contraseña generada a partir de la verificación en 2 pasos.
2. Haz clic en "Iniciar sesión" para autenticarte. Si las credenciales son correctas, se abrirá la ventana para enviar correos. Si la autenticación falla, se te mostrará un mensaje de error.
3. En “Desde:”, escribimos el mail de donde saldrán los envíos, si tenemos una cuenta vinculada a nuestro Gmail podremos usar esa como salida. 
4. Completamos los campos de Asunto y Cuerpo del mail.
5. Adjuntamos el Excel con los destinatarios. (Utilizar como modelo el archivo que se encuentra en la carpeta llamado “modelo_db_destinatarios”.
- Las columnas deben llamarse siempre ID, NOMBRE y CORREOS (CIA, NOMBRE y CORREOS en la app_v3). Los correos deben separarse por punto y coma. Ej.: correo1@gmail.com; correo2@gmail.com
6. Luego “Seleccionar carpetas CIAS” y seleccionamos la carpeta donde están los archivos Excel. Estos deben ser llamados 0002, 0016, etc.
7. Se puede tildar la opción de “Incluir nombre de la compañía en el asunto” que incluirá el respectivo nombre de cada CIA al final del asunto.

“Adjuntar y procesar archivo a separar (OPCIONAL)”: Se hace en un caso en que tengamos todos los registros juntos de las compañías y queremos que se separen por Numero de CIA.

7. Enviar

Tener en cuenta:
Los correos que contengan la letra “ñ” o tildes darán error en la ejecución, En caso de que esto suceda se deben eliminar o renombrar adecuadamente en el Excel de destinatarios y volver a cargarlo para continuar el envio de los mails. 
