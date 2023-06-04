import pandas as pd
import smtplib
import email.message
import time

vendas = pd.read_excel(r'C:\Users\Alexandre\Documents\Programacao\Python\Curso impressionador\vendas\Vendas.xlsx')

faturamento_loja = vendas.groupby('ID Loja')['Valor Final'].sum()
faturamento_loja =  faturamento_loja.reset_index()

quantidade_vendida_loja = vendas.groupby('ID Loja')['Quantidade'].sum()
quantidade_vendida_loja = quantidade_vendida_loja.reset_index()

ticket_medio = faturamento_loja['ID Loja']
ticket_medio = ticket_medio.reset_index()
ticket_medio['Ticket Medio'] = faturamento_loja['Valor Final'] / quantidade_vendida_loja['Quantidade']

df_total = faturamento_loja[['ID Loja', 'Valor Final']]
df_total['Quantidade'] = quantidade_vendida_loja['Quantidade']
df_total['Ticket Medio'] = ticket_medio['Ticket Medio']

nome = 'Alexandre'

faturamento_loja_html = df_total.to_html(index=False)

def enviar_email(lista_email):
    for emaill in lista_email:
        corpo_email = f"""
        <p>Segue a tabela com Faturamento de cada loja: </p>
        <br>
        {faturamento_loja_html}
        """

        msg = email.message.Message()
        msg['Subject'] = f"Indicadores Lojas"
        msg['From'] = 'SEU EMAIL AQUI'
        msg['To'] = emaill
        password = 'senha do app gmail aqui'
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email )

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Login Credentials for sending the mail
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print('Email enviado')
        time.sleep(2)


lista_emaill = ['lista de email aqui']
enviar_email(lista_emaill)