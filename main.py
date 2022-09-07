import pandas as pd
import re
# Regular expression for validating the mail
regex = re.compile(
    r"([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\"([]!#-[^-~ \t]|(\\[\t -~]))+\")@([-!#-'*+/-9=?A-Z^-~]+(\.[-!#-'*+/-9=?A-Z^-~]+)*|\[[\t -Z^-~]*])")

# Function to validate the mail using Regex module


def isValid(email):
    if re.fullmatch(regex, email):
        return email

    
    def sendEmail():
    try:
        listOfFinalEmails = readMails()
        print("Mail will be send of these emails:", listOfFinalEmails)
        # Reading message from HTML file
        HTMLFile = open("index.html", "r")
        messageHTML = HTMLFile.read()

        # SMTP Object
        print("In process...")
        mailServer = smt.SMTP("smtp.gmail.com", 587)
        mailServer.starttls()  # Starting the server

        # Setting up the email subject, to, from, and message
        fromEmail = senderMail
        toEmail = listOfFinalEmails
        Emailmessage = MIMEMultipart("alternative")
        Emailmessage['Subject'] = "Testing for Python Project"
        Emailmessage['from'] = senderMail

        # Login via email
        mailServer.login(senderMail, senderPass)
        print("Login Successful")

        # Attach the message and Multipart
        textMsg = MIMEText(messageHTML, "html")
        Emailmessage.attach(textMsg)

        # Sending the email
        mailServer.sendmail(fromEmail, toEmail, Emailmessage.as_string())
        mailServer.quit()
        print("sent")
        if(mailServer.send_message):
            return True
        else:
            return False

    except Exception as error:
        print(error)
