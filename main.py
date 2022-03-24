from sheet2api import Sheet2APIClient
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import date

client = Sheet2APIClient(api_url='https://sheet2api.com/v1/aRGKx6L8TRpm/livro1')
data_atual = date.today()
data_em_texto = '{}/{}/{}'.format(data_atual.day, data_atual.month,data_atual.year)

squad_result = {}

squads = [['Leandro'], ['Stefany', 'Marcos'], ['Marco', 'Ana'], ['Ruth', 'Gabriel'],
['Maykon', 'Maite'], ['Guilherme'], ['Dario', 'Giovana'], ['Leo', 'Lucas Ribas']]

date = data_em_texto

def get_status(worker, date):
    list_row = client.get_rows(query={'Colaborador': worker})
    list_dict = list_row[0]
    worker_status = list_dict[date]
    return worker_status
'''
for x in range(len(squads)):
    for y in range(len(squads[x])):
        #print(squads[x][y])
        #print(squads[x][y] + get_status(squads[x][y], '29/11/2021'))
        squad_result[squads[x][y]]= get_status(squads[x][y], data_em_texto)


print(squad_result)'''
squad_result = {'Leandro': 'HO', 'Stefany': 'HO', 'Marcos': 'HO', 'Marco': 'HO', 'Ana': 'Itb', 'Ruth': 'Itb', 'Gabriel': 'Itb', 'Maykon': 'Itb', 'Maite': 'Itb', 'Guilherme': 'Itb', 'Dario': 'Itb', 'Giovana': 'Itb', 'Leo': 'Itb', 'Lucas Ribas': 'Itb'}

def send_email(worker, date, list_dict):
    username = "le052437@intelbras.com.br"
    password = "Asadew12w1"
    mail_from = "le052437@intelbras.com.br"
    mail_to = "leonardo.lima@intelbras.com.br"
    mail_subject = "Controle de Escala"

    mimemsg = MIMEMultipart()
    mimemsg['From']=mail_from
    mimemsg['To']=mail_to
    mimemsg['Subject']=mail_subject

    variable_teste = "LEO LIMA"

    html = """\
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>Title</title>
</head>
<body>
<p>Bom dia,</p>
<p>Segue a seguir a escala programada da equipe para amanhã {date}:</p>
<h2>Controle de Escala</h2>
<p>Leandro: <b>{result_leandro}</b></p>
<p>Leo: <b>{result_leo}</b></p>
<p>Lucas Ribas: <b>{result_lucasribas}</b></p>
<h3>Squad Tico e Teco</h3>
<p>Stefany: <b>{result_stefany}</b></p>
<p>Marcos: <b>{result_marcos}</b></p>
<h3>Squad Unicornio</h3>
<p>Marco: <b>{result_marco}</b></p>
<p>Ana: <b>{result_ana}</b></p>
<h3>Squad Pra Frente</h3>
<p>Ruth: <b>{result_ruth}</b></p>
<p>Gabriel: <b>{result_Gabriel}</b></p>
<h3>Squad M&M</h3>
<p>Maykon: <b>{result_maykon}</b></p>
<p>Maite: <b>{result_maite}</b></p>
<p></p>
<p> Segue o link para acesso a planilha: https://onedrive.live.com/edit.aspx?cid=24ecdaf014356d2c&page=view&resid=24ECDAF014356D2C!58898&parId=24ECDAF014356D2C!101&app=Excel </p>
<p></p>
<p>Atenciosamente,</p>
<p>Leonardo Lima</p>
</body>
</html>
    """.format(date=data_em_texto,
               result_leandro=squad_result["Leandro"],
               result_leo=squad_result["Leo"],
               result_lucasribas=squad_result["Lucas Ribas"],
               result_stefany=squad_result["Stefany"],
               result_marcos=squad_result["Marcos"],
               result_marco=squad_result["Marco"],
               result_ana=squad_result["Ana"],
               result_ruth=squad_result["Ruth"],
               result_Gabriel=squad_result["Gabriel"],
               result_maykon=squad_result["Maykon"],
               result_maite=squad_result["Maite"]
                )

    # Turn these into plain/html MIMEText objects
    part2 = MIMEText(html, "html")
    mimemsg.attach(part2)


    #mimemsg.attach(MIMEText(mail_body, 'plain'))
    connection = smtplib.SMTP(host='smtp.office365.com', port=587)
    connection.starttls()
    connection.login(username,password)
    connection.send_message(mimemsg)
    connection.quit()

send_email('a','b','c')