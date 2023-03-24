import smtplib
import time
from datetime import datetime, timedelta
from email.message import EmailMessage
import emoji
import pandas as pd
import pyautogui as pg
import pyodbc
import pywhatkit
from termcolor import colored

email_subject = "Email test from Python" #Asunto del correo
sender_email_address = "##########@########.com"#Emisor del correo
receiver_email_address = "#######@######.com"#Receptor del correo, en este momento no se sabe quien sera el receptor
email_smtp = "smtp.office365.com"#servidor smtp
email_password = "########"

#variables que se cargarán en el correo y mensaje
abogado = None
juzgado = None
numeroExp = None
primerNombre = None
primerApellido = None
segundoApellido = None
identificacion = None
cuenta = None
nombreCompleto = None
completarJuzgado = None
nId = None
numeroRel = None
supervision = None

#Números a Notificar
numeroFase = #######
numeroProduccion = #######
numeroCiclos = ########
numeroMant = #######
mariana = #######
caro = #######
prueba = #######
supervisionNotificar = None
SupervisionCorreo = None
correoAbogado = None

#while True:
    #MonitoreoFechaAudiencia
try:
    connection = pyodbc.connect(
        'Driver={ODBC Driver 17 for SQL Server};'
        'Server=#.#.#.#;'
        "Database=###;"
        "UID=respaldo;"
        "PWD=1234;"
    )
    cursor = connection.cursor()
    cursor.execute("SELECT * FROM [AUDIENCIAS] WHERE Cast([FECHA AUDIENCIA] AS date) >= Cast( GETDATE() AS date) ")
    result = cursor.fetchall()
    time.sleep(2)
    print('Conexión Activa')
except Exception as e:
    print("Error al conectar a la base de datos")
    print(e)
    #break
if result == False:        
    #break
    print("Sin resultados en la consulta")
else :
    for x in result:

        #fechas para verificar
        fechaA = (x[1])
        abogado = (x[3])        
        fecha1 = (x[12])
        fecha2 = (x[13])
        fecha3 = (x[14])
        fechaActual = datetime.today()
        #contador de días restantes
        diasRestantes = (fechaA-fechaActual).days
        diasRestantes += 1;
        print(diasRestantes)


        #verificador de fecha de notificacion
        aviso1 = False
        if fecha1.year == fechaActual.year and fecha1.month == fechaActual.month and fecha1.day == fechaActual.day:
            aviso1 = True
        aviso2 = False
        if fecha2.year == fechaActual.year and fecha2.month == fechaActual.month and fecha2.day == fechaActual.day:
            aviso2 = True
        aviso3 = False
        if fecha3.year == fechaActual.year and fecha3.month == fechaActual.month and fecha3.day == fechaActual.day:
            aviso3 = True 


        if aviso1 or aviso2 or aviso3:
            #comienza asignación de valores a las variables
            nId = (x[0])          
            numeroRel = (x[16])
            numeroRel = numeroRel.strip()
            supervision = (x[2])
            supervision = supervision.strip()

            #######################contrato
            contrato = (x[4])
            if contrato == '-1':
                contratoLetras = "LISTO"
            else:
                contratoLetras = "PENDIENTE"
            #######################constanciaPagos
            constanciaPagos = (x[5])
            if constanciaPagos == '-1':
                constanciaPagosLetras = "LISTO"
            else:
                constanciaPagosLetras = "PENDIENTE" 
            #######################documentosAudiencia                                                         
            documentosAudiencia = (x[6])
            if documentosAudiencia == '-1':
                documentosAudienciaLetras = "LISTO"
            else:
                documentosAudienciaLetras = "PENDIENTE"
            #######################expedienteImpreso                        
            expedienteImpreso = (x[7])
            if expedienteImpreso == '-1':
                expedienteImpresoLetras = "LISTO"
            else:
                expedienteImpresoLetras = "PENDIENTE"
            #######################revisionExpediente                    
            revisionExpedientes = (x[8])
            if revisionExpedientes == '-1':
                revisionExpedientesLetras = "LISTO"
            else:
                revisionExpedientesLetras = "PENDIENTE"
            #######################alistarNegociacion                  
            alistarNegociacion = (x[9])
            if alistarNegociacion == '-1':
                alistarNegociacionLetras = "LISTO"
            else:
                alistarNegociacionLetras = "PENDIENTE"
            #######################consultarAbogado
            consultarAbogado = (x[10])
            if consultarAbogado == '-1':
                consultarAbogadoLetras = "LISTO"
            else:
                consultarAbogadoLetras = "PENDIENTE"  
            #######################revisionFinal
            revisionFinal = (x[11])
            if revisionFinal == '-1':
                revisionFinalLetras = "LISTO"
            else:
                revisionFinalLetras = "PENDIENTE"

            
            #Se realizan las consultas para extraer los datos faltantes para el correo
            cursor2 = connection.cursor()
            cursor2.execute("SELECT [ABOGADO], [JUZGADO], [NUMERO EXPEDIENTE] FROM [########] WHERE [ESTATUS EXPEDIENTE] = 'ACTIVO' or [ESTATUS EXPEDIENTE] = 'AUDIENCIA' and [NUMERO RELACIONAL] = ?",(numeroRel,))               
            result2 = cursor2.fetchall()
            cursor3 = connection.cursor()
            cursor3.execute("SELECT [PRIMER NOMBRE], [PRIMER APELLIDO], [SEGUNDO APELLIDO], [IDENTIFICACION], [CUENTA] FROM [#######] WHERE [NUMERO RELACIONAL] = ?",(numeroRel,))
            result3 = cursor3.fetchall()
            
            
            for y in result2:
                juzgado = (y[1])
                #print(str(juzgado))                
                numeroExp = (y[2])
                #print(str(numeroExp))

            for z in result3:
                primerNombre = (z[0])
                primerNombre = primerNombre.strip()
                primerApellido = (z[1])
                primerApellido = primerApellido.strip()
                segundoApellido = (z[2])
                segundoApellido = segundoApellido.strip()
                nombreCompleto = primerNombre + ' ' + primerApellido + ' ' + segundoApellido
                print(nombreCompleto)
                identificacion = (z[3])
                identificacion = identificacion.strip()
                cuenta = (z[4])

                cuenta = cuenta.strip()

            print(abogado)
            correoAbogado = "#####@#####.com" if abogado == "VM" else "#####@#####.com"
                nombreAbogado = "Abogado Goku" if abogado == "EO" else "Abogada Bulma" if abogado == "VM" else "abogado(a)"
            print(nombreAbogado)            
            
            if juzgado == 'GOICO 3':
                completarJuzgado = 'JUZGADO ESPECIALIZADO DE COBRO DEL II CIRCUITO JUDICIAL DE SAN JOSÉ, GOICOECHEA, SECCIÓN TERCERA'
            elif juzgado == 'GOICO 1':
                completarJuzgado = 'JUZGADO ESPECIALIZADO DE COBRO DEL II CIRCUITO JUDICIAL DE SAN JOSÉ, GOICOECHEA, SECCIÓN PRIMERA'
            elif juzgado == 'GOICO 2':
                completarJuzgado = 'JUZGADO ESPECIALIZADO DE COBRO DEL II CIRCUITO JUDICIAL DE SAN JOSÉ, GOICOECHEA, SECCIÓN SEGUNDA'
            elif juzgado == 'SAN CARLOS':
                completarJuzgado = 'JUZGADO ESPECIALIZADO DE COBRO DEL II CIRCUITO JUDICIAL DE ALAJUELA'
            elif juzgado == 'HEREDIA':
                completarJuzgado = 'JUZGADO ESPECIALIZADO DE COBRO DE HEREDIA'
            elif juzgado == 'CARTAGO':
                completarJuzgado = 'JUZGADO ESPECIALIZADO DE COBRO DE CARTAGO'
            elif juzgado == 'PUNTARENAS':
                completarJuzgado = 'JUZGADO ESPECIALIZADO DE COBRO DE PUNTARENAS'
            elif juzgado == 'LIMON':
                completarJuzgado = 'JUZGADO ESPECIALIZADO DE COBRO DEL I CIRCUITO JUDICIAL DE LA ZONA ATLÁNTICA'
            elif juzgado == 'SAN RAMON':
                completarJuzgado = 'JUZGADO DE COBRO DEL III CIRCUITO JUDICIAL DE ALAJUELA (SAN RAMÓN)'
            elif juzgado == 'PRIMERO':
                completarJuzgado = 'JUZGADO PRIMERO COBRATORIO'
            elif juzgado == 'SEGUNDO':
                completarJuzgado = 'JUZGADO SEGUNDO COBRATORIO'
            elif juzgado == 'TERCERO':
                completarJuzgado = 'JUZGADO TERCERO COBRATORIO'
            elif juzgado == 'ALAJUELA':
                completarJuzgado = 'JUZGADO ESPECIALIZADO DE COBRO DEL I CIRCUITO JUDICIAL DE ALAJUELA'
            elif juzgado == 'GRECIA':
                completarJuzgado = 'JUZGADO ESPECIALIZADO DE COBRO DE GRECIA'
            elif juzgado == 'LIBERIA':
                completarJuzgado = 'JUZGADO DE COBRO DEL I CIRCUITO JUDICIAL DE GUANACASTE (LIBERIA)'
            elif juzgado == 'SANTA CRUZ':
                completarJuzgado = 'JUZGADO DE COBRO DEL II CIRCUITO JUDICIAL DE GUANACASTE (SANTA CRUZ)'
            elif juzgado == 'POCOCI':
                completarJuzgado = 'JUZGADO ESPECIALIZADO DE COBRO DE POCOCÍ'
            elif juzgado == 'GOLFITO':
                completarJuzgado = 'JUZGADO ESPECIALIZADO DE COBRO DE GOLFITO'
            elif juzgado == 'PEREZ ZELEDON':
                completarJuzgado = 'JUZGADO DE COBRO DEL I CIRCUITO JUDICIAL ZONA SUR (PÉREZ ZELEDÓN)'
            

            #definición de variable para las fechas
            fechaAudienciaFormatoCorto = fechaA.strftime('%d/%m/%y')
            # Obtener el nombre del día de la semana y el mes en formato completo y traducido
            dia_semana = fechaA.strftime("%A").replace("Monday", "lunes").replace("Tuesday", "martes").replace("Wednesday", "miércoles").replace("Thursday", "jueves").replace("Friday", "viernes").replace("Saturday", "sábado").replace("Sunday", "domingo")
            mes = fechaA.strftime("%B").replace("January", "enero").replace("February", "febrero").replace("March", "marzo").replace("April", "abril").replace("May", "mayo").replace("June", "junio").replace("July", "julio").replace("August", "agosto").replace("September", "septiembre").replace("October", "octubre").replace("November", "noviembre").replace("December", "diciembre")
            # Obtener la hora en formato de 12 horas y la indicación AM/PM
            hora = fechaA.strftime("%I:%M %p")
            time.sleep(2)#Toma un receso de 2 segundos
            

            #configurando los correos para el envío
            print(colored('Datos Obtenidos:', 'blue', 'on_white'))
            to_addresses = ["########@######.com"]
            to_addresses = ["########@######.com", "########@######.com"]
            to_addresses = []
            if supervision == "CICLOS":
                SupervisionCorreo = "########@######.com"
                print(SupervisionCorreo)
                to_addresses.append(SupervisionCorreo)
            elif supervision == "PRODUCCION":
                SupervisionCorreo = "########@######.com"
                print(SupervisionCorreo)
                to_addresses.append(SupervisionCorreo)
            elif supervision == "FASE":
                SupervisionCorreo = "########@######.com"
                print(SupervisionCorreo)
                to_addresses.append(SupervisionCorreo)
            elif supervision == "MANTENIMIENTO":
                SupervisionCorreo = "########@######.com"
                print(SupervisionCorreo)
                to_addresses.append(SupervisionCorreo)
            to_addresses.append(correoAbogado)
        


            email_subject = (emoji.emojize(":warning:")) +  (" Notificación Audiencia ")  + (emoji.emojize(":warning:"))
            # Create an email message object
            message = EmailMessage() #Se crea un objeto de la clase EmailMessage(), para configurar y consultar las cabeceras
            message['Subject'] = email_subject
            message['From'] = sender_email_address
            message['To'] = ", ".join(to_addresses)
            html = ('''
<!DOCTYPE html>
<html>
<head>
    <title>CORREO</title>
</head>    
    <body style=" width:350px; background-color:#D4E6F1; border: 6px solid black; border-radius: 20px; display: flex; align-items: center; word-break:break-all; margin: 0 0 1em 1em;">
        <div style="padding:20px 0px; ">
            <div style= "background-color:#C70039;  width:275px; padding:10px;  float: left; display: flex; align-items: center; word-break:break-all; margin: 0 0 1em 1em; box-shadow: 10px 5px 5px black; border-radius: 30px; ">
                <div><img src="http://186.26.125.218/mendez.jpeg" style="height: 65px;  border: 2px solid white; border-radius: 20px;">  &nbsp;  &nbsp; <img src="http://186.26.125.218/Poder/Audiencia2.gif" style= "height: 20px;"></div></div>

            <div  style="background-color:#bdc700; padding:10px 20px;  float: left; display: flex; align-items: center;word-break:break-all; margin: 0 0 1em 1em; color:#f13000; box-shadow: 10px 5px 5px black;">
                <h2 style="background-color:#bdc700; padding:10px 20px; ;"> Audiencia : {fechaAudienciaFormatoCorto} &nbsp; &nbsp;  {nombreCompleto}</h2>
            </div>
        
                        <div style="background-color:#13b469; color:white; width:275px; padding:5px 10px;  float: left; display: flex; align-items: center; word-break:break-all; margin: 0 0 1em 1em; box-shadow: 2px 2px 2px 1px rgba(0, 0, 0, 1.2);">    
                <ul style="background-color:#0078D4;  box-shadow: 10px 5px 5px black;">
                    <h2> Datos del Demandado: &nbsp;  &nbsp; &nbsp;  </h2> 
                
                    <li style="background-color:#075577;padding:1px 1px; ">Abogado: <h3>{nombreAbogado}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px;">Juzgado:<h3>{completarJuzgado}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px;">Expediente:<h3>{numeroExp}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px;">Día de la Audiencia:<h3>{dia_semana} {fechaA.day} de {mes} {fechaA.year} a las {hora}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px;">Nombre:<h3>{nombreCompleto}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px;">Cédula:<h3>{identificacion}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px;">Tarjeta:<h3>{cuenta}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px;">Cuenta:<h3>{cuenta}</h3></li>
                </ul> 
                    </div>
    
                        <div style="background-color:#13b469; color:white; width:275px; padding:5px 10px;   float: left; display: flex; align-items: center; word-break:break-all; margin: 0 0 1em 1em; box-shadow: 2px 2px 2px 1px rgba(0, 0, 0, 1.2);">    
                <ul style="background-color:#0078D4; ; box-shadow: 10px 5px 5px black; ">
                    <h2> Comprobaciones &nbsp;  &nbsp; &nbsp;  </h2> 
                    <li style="background-color:#075577;padding:1px 1px; ">Solicitar Contrato:<h3>{contratoLetras}</h3></li>                        
                    <li style="background-color:#075577;padding:1px 1px; ">Solicitar Constancia de Pagos:<h3>{constanciaPagosLetras}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px; ">Alistar Documentos de Audiencia:<h3>{documentosAudienciaLetras}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px; ">Expediente Impreso:<h3>{expedienteImpresoLetras}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px; ">Revisión de Expediente:<h3>{revisionExpedientesLetras}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px; ">Alistar Negociaciones:<h3>{alistarNegociacionLetras}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px; ">Consultar Don Erick:<h3>{consultarAbogadoLetras}</h3></li>
                    <li style="background-color:#075577;padding:1px 1px; ">Revisión Final: <h3>{revisionFinalLetras}</h3></li>
                </ul> 
                    </div>
        </div>
</body>
</html>
            ''').format(**locals())#.format() nos permite incluir en una cadena de texto, variables que se encuentra en el script, usando {}
            #locals()retorna un diccionario con todas las variable y simbolos del actual programa.
            message.set_content(html, subtype='html')#En el contenido del mensaje de correo se adjunta la variable html de subtipo html
            #message.set_content(mensajes)
            server = smtplib.SMTP(email_smtp, '587')#smtplib define un objeto de sesión de cliente SMTP, con el nombre del host y el puerto
            server.ehlo()#Envia un eco antes de enviar el correo
            server.starttls()#Se pone la conexión SMTP en modo TLS(Transport Layer Security)
            server.login(sender_email_address, email_password)#Loguearse con la direccion del receptor y las contraseña
            server.send_message(message)#Se envia el mensaje 
            server.quit()#Se cierra la conexión
            
            
            #envio de mensajes por whatsapp
            print("Va empezar la pausa de 10 segundos")
            time.sleep(8)
            #combo = [prueba]
            combo = [mariana, caro, prueba]
            if supervision == "CICLOS":
                supervisionNotificar = numeroCiclos
                combo.append(supervisionNotificar)
                print(supervisionNotificar)
            elif supervision == "PRODUCCION":
                supervisionNotificar = numeroProduccion
                combo.append(supervisionNotificar)
                print(supervisionNotificar)                
            elif supervision == "FASE":
                supervisionNotificar = numeroFase
                combo.append(supervisionNotificar)
                print(supervisionNotificar)                
            elif supervision == "MANTENIMIENTO":
                supervisionNotificar = numeroMant
                combo.append(supervisionNotificar)
                print(supervisionNotificar)                                
            first = True
            for celular in combo:
                hora = (datetime.now().strftime('%H'))
                minr = (datetime.now().strftime('%M'))
                min = int(minr)+2
                if min >= 60:
                    min = 0
                    hora = int(hora)+1
                print ("Empezando a recorrer arreglo de numeros")
                mensaje = (emoji.emojize(":warning:")) + ('*Audiencia*') + (emoji.emojize(":warning:")) + ('Faltan ') + str(diasRestantes) + (' días para la Audiencia con el demandado: ') + str(nombreCompleto) + ('\n') + ('*Expediente:*') + ('\n') + str(numeroExp) + ('\n') + ('Cédula: ') + ('\n') + str(identificacion) + ('\n') + ("Tarjeta: ") + str(cuenta) + ('\n') + ("Cuenta: ") + str(cuenta) + ('\n') + ("Revisar los documentos pendientes para la Audiencia")
                time.sleep(2)
                try:
                    pywhatkit.sendwhatmsg("+506"+str(celular), str(mensaje),int(hora), int(min))
                    print(colored('Mensaje ENVIADO', 'yellow', 'on_red'))
                    if first:
                        time.sleep(2)
                        first=False
                    time.sleep(4)
                    pg.hotkey('ctrl', 'w')
                    time.sleep(2)
                    pg.press('enter')
                    #time.sleep(2)
                except :
                    print(colored('ERROR de Whatsapp', 'yellow', 'on_red'))
                    #break
