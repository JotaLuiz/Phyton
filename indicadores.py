# ETO 2.0
# Para adicionar uma nova opção de envio é necessário a criação de um novo parâmetro de "body","arc" e "sub"
# e atribuição a 1uma variável

import smtplib
import pwinput


from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ------------------------------------------------------------------------------------
# Intro + Variáveis

regs = []
emails = []
with open('Dados.txt', mode='r', encoding='utf-8') as contacts_file:
    for a_contact in contacts_file:
        regs.append(a_contact.split()[0])
        emails.append(a_contact.split()[1])

frm = str(input('Digite o e-mail de envio: '))
psswrd = pwinput.pwinput('Digite a senha do e-mail: ')
nam = str(input('Digite seu nome para a assinatura: '))
datname = str(input('Digite mês para corpo do e-mail e assunto: '))
anodat = str(input('Digite ano para corpo do e-mail e assunto: '))
diadat = str(input('Digite data de hoje para assunto: '))
item = str(input('Digite bom dia/boa tarde: '))
print('Você digitou: 'f'{item}')
print(''' 
1. Planilha de Indicadores
2. Planilha de Estoque
''')
out = int(input('Selecione uma opção: '))

# ------------------------------------------------------------------------------------
# Classe


class SendDis:
    def __init__(self, body, arc, sub):
        self.body = body
        self.arc = arc
        self.sub = sub

        pass

    def send(self):
        for reg, email in zip(regs, emails):
            try:
                fromaddr = frm
                msg = MIMEMultipart()
                receivers = [email]

                msg['From'] = fromaddr
                msg['To'] = ', '.join(receivers)
                msg['Subject'] = self.sub

                msg.attach(MIMEText(self.body, 'plain'))

                filename = f'{reg}{self.arc}'

                attachment = open(filename, 'rb')

                part = MIMEBase('application', 'octet-stream')
                part.set_payload((attachment).read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition',
                                "attachment; filename= %s" % filename)

                msg.attach(part)

                attachment.close()

                server = smtplib.SMTP('smtp.outlook.com', 587)
                server.starttls()
                server.login(fromaddr, psswrd)
                text = msg.as_string()
                server.sendmail(fromaddr, receivers, text)
                server.quit()
                print(f'\nEmail enviado para a regional {reg} com sucesso!')
            except:
                print(f'Falha ao enviar email para a regional {reg}.')
                print(' ')

# ------------------------------------------------------------------------------------
# Parâmetros


# Indicadores
ind_body = f'''\nPrezados, {item}!

Seguem indicadores de Gramagem, Sobra Limpa e Resto ingesta consolidados de {datname}.

Ressaltamos a importância do acompanhamento dos dados para buscarmos melhoria constante em nosso Custo Alimentar e Satisfação. 

Qualquer dúvida estou à disposição!

att,

{nam}
'''
ind_arc = '.xlsx'
ind_sub = 'Planilha de Gramagem - ' f'{datname}/'f'{anodat} - 'f'{diadat}'

# Estoque
est_body = f'''\nPrezados, {item}!

Segue o indicador atualizado do estoque sistêmico enviado semanalmente.

Ressalto que esta posição reflete o estoque sistêmico desta madrugada, com custos com impostos e logística.

Qualquer dúvida estou à disposição!

att,

{nam}
'''
est_arc = '.xlsx'
est_sub = 'Planilha de Estoque - ' f'{datname}/'f'{anodat} - 'f'{diadat}'

# ------------------------------------------------------------------------------------
# Definição

ind = SendDis(ind_body, ind_arc, ind_sub)
est = SendDis(est_body, est_arc, est_sub)

# ------------------------------------------------------------------------------------
# Execução

if out == 1:
    x = str(input('''\nVocê selecionou a opção "Planilha de Indicadores", continuar?
Sim/Não: '''))
    if x.lower() == 'sim':
        ind.send()
    else:
        print('Programa encerrado.')
if out == 2:
    x = str(input('''\nVocê selecionou a opção "Planilha de Estoque", continuar?
Sim/Não: '''))
    if x.lower() == 'sim':
        est.send()
    else:
        print('Programa encerrado.')        
        
    input("Pressione <enter> para encerrar!")