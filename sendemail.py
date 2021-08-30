import smtplib
import imghdr
from email.message import EmailMessage


def sendmail(receiver_email, image_name, EMAIL_ADDRESS, APP_PASSWORD, SUBJECT, BODY):
    msg = EmailMessage()
    msg['Subject'] = SUBJECT
    msg['From'] = EMAIL_ADDRESS
    msg.set_content(BODY)

    RECEIVER_EMAIL = receiver_email
    IMAGE_NAME = image_name
    msg['To'] = RECEIVER_EMAIL
    with open(IMAGE_NAME, 'rb') as f:
        image_data = f.read()
        image_type = imghdr.what(f.name)
        msg.add_attachment(image_data, maintype='image', subtype=image_type, filename=image_name)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(EMAIL_ADDRESS, APP_PASSWORD)
        server.send_message(msg)
        server.quit()
        del msg