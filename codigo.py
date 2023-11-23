import os 
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

caminho = "bases"
arquivos = os.listdir(caminho)

tabela_consolidada = pd.DataFrame()

for nome_arquivo in arquivos:
    tabela_vendas = pd.read_csv(os.path.join(caminho, nome_arquivo))
    tabela_vendas ["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(tabela_vendas["Data de Venda"],
                                                                                        unit="d")
    tabela_consolidada = pd.concat([tabela_consolidada, tabela_vendas])

tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda")
tabela_consolidada = tabela_consolidada.reset_index(drop=True)
tabela_consolidada.to_excel("Vendas.xlsx", index=False)


# Envio automatico pelo Gmail
# Configurações de e-mail
de = 'leonardorfragoso@gmail.com'
para = 'leonardorfragoso2@gmail.com'
senha = 'ppwbgavdaaaowxeo'
assunto = 'Bases compiladas'
mensagem = 'Automação para Compilar bases de dados em uma unica planila de Excel e enviar automaticamente por email.'

# Configuração do servidor SMTP do Gmail
servidor_smtp = 'smtp.gmail.com'
porta = 587

# Crie uma mensagem de e-mail
mensagem_email = MIMEMultipart()
mensagem_email['From'] = de
mensagem_email['To'] = para
mensagem_email['Subject'] = assunto
mensagem_email.attach(MIMEText(mensagem, 'plain'))

# Anexando um arquivos
caminho = os.getcwd()
arquivo_anexo = os.path.join(caminho, "Vendas.xlsx")
with open(arquivo_anexo, 'rb') as arquivo:
    part = MIMEApplication(arquivo.read(), _subtype="pdf")
part.add_header('Content-Disposition', 'attachment', filename=arquivo_anexo)
mensagem_email.attach(part)

# Conectando ao servidor SMTP e enviando o e-mail
try:
    servidor = smtplib.SMTP(servidor_smtp, porta)
    servidor.starttls()
    servidor.login(de, senha)
    servidor.sendmail(de, para, mensagem_email.as_string())
    servidor.quit()
    print('E-mail enviado com sucesso')
except Exception as e:
    print('Erro ao enviar o e-mail:', str(e))