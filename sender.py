from mailer import Mailer


def sender(ereceiver,massage):
    mzil = Mailer(email="", password="")
    mzil.send(receiver=ereceiver,
            subject="UHB Custody",
            message=massage )
    



