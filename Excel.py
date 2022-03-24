import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

date_today = datetime.today().strftime('%d/%m/%Y')

df_escala = pd.read_excel('Escala.xlsx', index_col=0)
workers = df_escala.loc['Marcos'][date_today]

dict_schedule = {}

for colaborador in df_escala.index:
    status = df_escala.loc[colaborador][date_today]
    dict_schedule[colaborador] = status



def send_email(dict_escala, date_to_insert):
    username = "le052437@intelbras.com.br"
    password = "Asadew12w1"
    mail_from = "le052437@intelbras.com.br"
    mail_to = "leonardo.lima@intelbras.com.br"
    mail_subject = "Controle de Escala"

    mimemsg = MIMEMultipart()
    mimemsg['From']=mail_from
    mimemsg['To']=mail_to
    mimemsg['Subject']=mail_subject

    html = """\
    <html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
    <head>
        <title>Title</title>
    </head>
    <body>
    <p>Bom dia,</p>
    <p>Segue a seguir a escala programada da equipe para hoje {date}:</p>
    <h2>Controle de Escala</h2>
    <p>{escala}</p>
    <p></p>
    <p> Segue o link para acesso a planilha: https://onedrive.live.com/edit.aspx?cid=24ecdaf014356d2c&page=view&resid=24ECDAF014356D2C!58898&parId=24ECDAF014356D2C!101&app=Excel </p>
    <p></p>
    <p>Atenciosamente,</p>
    <p>Leonardo Lima</p>
    </body>
    </html>
        """.format(date=date_to_insert,
                   escala=dict_escala
                   )

    # Turn these into plain/html MIMEText objects
    part2 = MIMEText(html, "html")
    mimemsg.attach(part2)

    # mimemsg.attach(MIMEText(mail_body, 'plain'))
    connection = smtplib.SMTP(host='smtp.office365.com', port=587)
    connection.starttls()
    connection.login(username, password)
    connection.send_message(mimemsg)
    connection.quit()


send_email(dict_schedule, date_today)