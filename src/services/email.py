import smtplib
import ssl
from email.message import EmailMessage


class EmailService:
    def __init__(self):
        self.sendConfirmationMessagefilesCreated = []
        self.filesUpdated = []
        self.filesCreated = []
        self.email_sender = "aldocitalanortiz@gmail.com"
        self.email_password = "sumpbqklblqdpryn"
        self.email_receiver = "aldocitalan81@gmail.com"

    def sendConfirmationMessage(self):
        message = EmailMessage()
        message["Subject"] = "Archivos creados y actualizados"
        message["From"] = self.email_sender
        message["To"] = self.email_receiver

        body = f""" 

            <h1 style = 'color: black;'>Nuevos archivos generados</h1>

            <p>Los archivos se encuentran en la siguiente ruta: </p>
            <br>
            Archivos creados: {' ,'.join(self.filesCreated)}
            <br>
            Archivos actualizados: {' ,'.join(self.filesUpdated)}

            

       """
        message.set_content(body, subtype="html")

        context = ssl.create_default_context()

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(self.email_sender, self.email_password)
            server.sendmail(self.email_sender, self.email_receiver, message.as_string())
