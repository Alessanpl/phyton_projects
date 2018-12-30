# encoding: utf-8

"""
  
  Cobra clientes com pagamento atrasado por meio de mensagem automática por e-mail

  1. Busca BD SQLLite por devedores empregando SQL
  2. Calcula quanto tempo está atrasado
  3. Monta mensagem para enviar por e-mail
  4. Envia mensagem por e-mail

"""

# Modulo para manipulação de email
import os, smtplib, sqlite3, datetime
from email.mime.text import MIMEText

# localizacao do banco de dados
pasta = "/Users/Alessanpl 1/Google Drive/Banco de Dados Bento Raro/"

if not os.path.exists(pasta + 'raroaulas.db'):
  print('O banco de dados ' + 'raroaulas.db' + ' não foi encontrado!')

# Envia e-mail simples
def EnviaEmailSimples(assunto, de, para, mensagem):
    msg = MIMEText(mensagem)
    msg['Subject'] = assunto
    msg['From'] = de
    msg['To'] = para

    # Enviando o e-mail
    Consmtp.send_message(msg)

# Logando no servidor do GMAIL
Consmtp = smtplib.SMTP("smtp.gmail.com", 587)
Consmtp.starttls()
Consmtp.login("alessanpl@gmail.com", "alessandro110")

# Salvar no SQLlite 
con = sqlite3.connect(pasta + 'raroaulas.db')

# Criando um cursor
cur = con.cursor()

# Executa a busca no banco
cur.execute("SELECT * FROM raroaulas WHERE nf_paga = '0' and data_prev is NULL")

# Calculo da quantidade de dias
def QtdDiasEntreDatas(Data_ini, Data_fin):
    d2 = datetime.datetime.strptime(Data_ini, '%d/%m/%Y')
    d1 = datetime.datetime.strptime(Data_fin, '%d/%m/%Y')

    return abs((d2 - d1).days)

for row in cur.fetchall(): 
    nf            = row[0]  
    data          = row[1]
    diarias_pagas = row[6]
    modulo        = row[9]
    e_mail_fin    = row[11]
    resp_fin      = row[12].split()[0]
    
    data_atual = datetime.datetime.today().strftime('%d/%m/%Y')
    quant_dias = QtdDiasEntreDatas(data_atual, data) 
    
    mensagem = """
      Olá {},

         Por gentileza, peço que verifique e me retorne a respeito do status de minha 
         nota fiscal {} emitida pela RARO PROJECT TREINAMENTOS LTDA em {} ({} dias atrás)
         {}referente as aulas do(s) módulo(s) {} 
         no MBA de Gerenciamento de Projetos.

      Grato,   
      Alessandro Prudêncio Lukosevicius, 
      Professor, Consultor, Pesquisador e Escritor,
      Doutor, PMP, PRINCE2 Approved Trainer, MSP Practitioner\n\n""" 

    
    if diarias_pagas == '0':
       mensagem = mensagem.format(resp_fin, nf, data, quant_dias, ' e das minhas diárias ', modulo)
    else:   
       mensagem = mensagem.format(resp_fin, nf, data, quant_dias, ' ', modulo)

    EnviaEmailSimples("Status de NF", "alessanpl@gmail.com", e_mail_fin, mensagem)
    print('Mensagem enviada para {} com e-mail {}'.format(resp_fin, e_mail_fin))

Consmtp.quit()
cur.close()
con.close()