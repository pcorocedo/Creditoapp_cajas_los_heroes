
import win32com.client
import os
import pandas as pd

def config (variable):
    """
    #Consulta de variable del libro de excel "Data/Config.xlsx"
    #ejemplo solo config(el nombre de la variable)
    """
    configuracion = pd.read_excel("Data/Config.xlsx",sheet_name='Variables ')
    (columna , celdas) = configuracion.shape
    x =-1
    while x < int(columna):
        x = 1+x
        if configuracion.iloc[int(x),0] == variable : break
    return configuracion.iloc[int(x),(1)]


def leer_email():
    """
    Lee los email y guarda los adjuntos
    :return:
    """
    try:
        # Carpeta de destino para guardar los adjuntos
        dest_folder = str(config("inputPL"))

        # Conexión a Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Selección de la bandeja de entrada
        inbox = outlook.GetDefaultFolder(6)

        # Búsqueda de todos los correos electrónicos no leídos
        messages = inbox.Items.Restrict("[Unread]=true")


        for msg in messages:
            # Obtener los datos del mensaje
            subject = msg.Subject
            sender = msg.SenderEmailAddress
            adjunto =msg.Attachments
            print("el adjunto es ",adjunto)
            print(f"Mensaje no leído: {subject} de {sender}")

            # Marcar el mensaje como leído
            msg.UnRead = False

            # Recorrer los adjuntos del mensaje
            for attachment in msg.Attachments:
                # Guardar el adjunto en la carpeta de destino
                attachment.SaveAsFile(os.path.join(dest_folder, attachment.FileName))

        print("Descarga de adjuntos completada.")
    except:
        print("Error en la descarga de adjuntos completada.")


