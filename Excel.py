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

print(dict_schedule)


def send_email(dict_escala, date_to_insert):
    username = "le052437@intelbras.com.br"
    password = "Asadew12w1"
    mail_from = "le052437@intelbras.com.br"
    mail_to = "," .join(["leonardo.lima@intelbras.com.br", "limarodriguesleonardo@gmail.com"])
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
    <table border="1">
    <tr>
        <td> Nome </td>
        <td> Escala </td>        
    </tr>
    <tr>
        <td> {colaborador1} </td>
        <td> {status_colaborador1} </td>        
    </tr>
    <tr>
        <td> {colaborador2} </td>
        <td> {status_colaborador2} </td>        
    </tr>
    <tr>
        <td> {colaborador3} </td>
        <td> {status_colaborador3} </td>        
    </tr>
    <tr>
        <td> {colaborador4} </td>
        <td> {status_colaborador4} </td>        
    </tr>
    <tr>
        <td> {colaborador5} </td>
        <td> {status_colaborador5} </td>        
    </tr>
    <tr>
        <td> {colaborador6} </td>
        <td> {status_colaborador6} </td>        
    </tr>
    <tr>
        <td> {colaborador7} </td>
        <td> {status_colaborador7} </td>        
    </tr>
    <tr>
        <td> {colaborador8} </td>
        <td> {status_colaborador8} </td>        
    </tr>
    <tr>
        <td> {colaborador9} </td>
        <td> {status_colaborador9} </td>        
    </tr>
    <tr>
        <td> {colaborador10} </td>
        <td> {status_colaborador10} </td>        
    </tr>
    <tr>
        <td> {colaborador11} </td>
        <td> {status_colaborador11} </td>        
    </tr>
    <tr>
        <td> {colaborador12} </td>
        <td> {status_colaborador12} </td>        
    </tr>
    <tr>
        <td> {colaborador13} </td>
        <td> {status_colaborador13} </td>        
    </tr>
    <tr>
        <td> {colaborador14} </td>
        <td> {status_colaborador14} </td>        
    </tr>  
    </table>    
    <p></p>
    <p> Segue o link para acesso a planilha: https://onedrive.live.com/edit.aspx?cid=24ecdaf014356d2c&page=view&resid=24ECDAF014356D2C!58898&parId=24ECDAF014356D2C!101&app=Excel </p>
    <p></p>
    <p>Atenciosamente,</p>
    <p>Leonardo Lima</p>
    </body>
    </html>
        """.format(date=date_to_insert,
                   escala=dict_escala,
                   colaborador1="Stefany",
                   status_colaborador1=dict_schedule["Stefany"],
                   colaborador2="Marcos",
                   status_colaborador2=dict_schedule["Marcos"],
                   colaborador3="Ana",
                   status_colaborador3=dict_schedule["Ana"],
                   colaborador4="Mendonça",
                   status_colaborador4=dict_schedule["Mendonça"],
                   colaborador5="Ruth",
                   status_colaborador5=dict_schedule["Ruth"],
                   colaborador6="Gabriel",
                   status_colaborador6=dict_schedule["Gabriel"],
                   colaborador7="Gustavo",
                   status_colaborador7=dict_schedule["Gustavo"],
                   colaborador8="Maykon",
                   status_colaborador8=dict_schedule["Maykon"],
                   colaborador9="Maitê",
                   status_colaborador9=dict_schedule["Maitê"],
                   colaborador10="Guilherme",
                   status_colaborador10=dict_schedule["Guilherme"],
                   colaborador11="Elionai",
                   status_colaborador11=dict_schedule["Elionai"],
                   colaborador12="Priscila",
                   status_colaborador12=dict_schedule["Priscila"],
                   colaborador13="Leo",
                   status_colaborador13=dict_schedule["Leo"],
                   colaborador14="Lucas Ribas",
                   status_colaborador14=dict_schedule["Lucas Ribas"],
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