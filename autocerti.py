import cv2
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import stdiomask

# assigning some constants
font = cv2.FONT_HERSHEY_TRIPLEX
scale = 0.9
thickness = 1

# path where generated certificates will be saved
path = r"L:\\autoCerti\certificates\\"

# reading the excel sheet containing name and email id of students.
data = pd.read_excel(r"L:\autoCerti\test.xls", sheet_name="Sheet1")
names = list(data['Name'])
receivers = list(data['email id'])

# taking email crendentials of sender
sender = stdiomask.getpass(prompt="enter Email Id:")
password = stdiomask.getpass(prompt='Enter your password:')

# reading the test certificate template
img = cv2.imread(r"L:\autoCerti\test.jpg")
img = cv2.resize(img, (512, 512))

# taking the start point and end point of baseline over which text will be written
coordX = list()
coordY = list()


def getCoords(event, x, y, flags, param):
    if event == cv2.EVENT_LBUTTONDOWN:
        coordX.append(x)
        coordY.append(y)


cv2.imshow("SELECT START AND END POINTS OF BASELINE", img)
cv2.setMouseCallback("SELECT START AND END POINTS OF BASELINE", getCoords)
cv2.waitKey(0)
cv2.destroyAllWindows()


# Generate certificate and mail to respective receiver
def generateCerti(names, img, coordX, coordY,):

    for text in names:
        img = cv2.imread(r"L:\autoCerti\test.jpg")
        img = cv2.resize(img, (512, 512))
        img1 = img
        tsize = cv2.getTextSize(text, font, scale, thickness)

        baseLen = coordX[1]-coordX[0]
        if baseLen >= tsize[0][0]:
            ext = baseLen-int(tsize[0][0])
        else:
            ext = int(tsize[0][0])-baseLen
        print(text)
        name = ("certificate_{}.png").format(text)
        filename = path + name
        X, Y = coordX[0]-ext//2, coordY[0]-tsize[1]
        output = cv2.putText(img1, text, (X, Y),
                             fontFace=cv2.FONT_HERSHEY_TRIPLEX, fontScale=scale, color=(255, 0, 0), thickness=thickness)

        cv2.imwrite(filename, output)
        print("certification is completed !!!")

        # Sending email to receiver
        emailGen(filename, name, sender, password,
                 receiver=receivers[names.index(text)], text=text)

# Generate Email with  a subject , header , body and certificate as attachment


def emailGen(filename, name, sender, password, receiver, text):

    subject = "Certification of successful Testing"
    body = ("Dear {},\nCongratualations !!!\nWe have successfully tested the autoCerti Programme.\nThank you for being a part of this.\nBest Regards\nJay").format(text)

    msg = MIMEMultipart()
    password = password
    msg['From'] = sender
    msg['To'] = receiver
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    filename
    attachment = open(filename, 'rb')

    part = MIMEBase('image', 'png', filename=name)

    part.add_header('Content-Disposition', 'attachment', filename=name)
    part.add_header('X-Attachment-Id', '0')
    part.add_header('Content-ID', '<0>')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    msg.attach(part)
    text = msg.as_string()

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(sender, password)

    server.sendmail(sender, receiver, text)
    print("certificate is successfully mailed !!!")

    server.quit()


generateCerti(names, img, coordX, coordY,)
