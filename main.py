# importar a base de dados com a biblioteca Pandas e o openpyxl
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

tabela_vendas = pd.read_excel('Vendas.xlsx')


# visualizar a base de dados
#pd.set_option('display.max_columns', None)
#print(tabela_vendas)

# faturamento por loja
def  get_total_value():

    faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

    return faturamento

#print(get_total_value())

# quantidade de produtos vendidos por loja
def get_sales_products():

    quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()

    return quantidade

#print(get_sales_products())

# ticket medio por produto em cada loja
def medium_ticket(faturamento, quantidade):
    ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
    ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
    return ticket_medio

print(medium_ticket(get_total_value(),get_sales_products()))

# enviar em email com relatório

def send_email(faturamento,quantidade,ticket_medio):
    # Defina as credenciais e o servidor SMTP
    smtp_server = "smtp.gmail.com"  # Por exemplo, para Gmail
    smtp_port = 587  # Para Gmail, a porta padrão para TLS é 587
    sender_email = "mateusluizgonzaga@gmail.com"
    receiver_email = "mateusluizgonzaga@gmail.com"
    password = ""  # Ou melhor, use um app password no caso do Gmail

    # Criando o e-mail
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = 'Relatório de vendas por Loja'
    body = f'''
    <h3>Prezados,</h3>
        

        <p>Segue o Relatório de vendas por cada loja.</p>

        <p>Faturamento:</p>
            {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

        <p>Quantidade Vendida:</p>
            {quantidade.to_html()}

        <p>Ticket Médio dos produtos de cada Loja:</p>
            {ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

        <p>Qualquer duvida estou a disposição.</p>

        <p>Att.,</p>
        <p>Mateus</p>
        '''
    msg.attach(MIMEText(body, 'html'))

    # Conectando ao servidor SMTP e enviando o e-mail
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Inicia a criptografia
            server.login(sender_email, password)  # Faz login
            text = msg.as_string()  # Converte a mensagem para string
            server.sendmail(sender_email, receiver_email, text)
            print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")



send_email(get_total_value(), get_sales_products(),medium_ticket(get_total_value(), get_sales_products()))


