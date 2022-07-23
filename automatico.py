#importanto as bibliotecas necessárias
import pandas as pd
import win32com.client as win32

#diretório do arquivo em excel
arq_excel = pd.read_excel(r'diretorio\ondeoarquivo\esta.xlsx')

#percorendo a lista de nomes
for i, email in enumerate(arq_excel['E-mail']):
    nomes = arq_excel.loc[i, 'Nome']
    
    #criando email
    mail = outlook.CreateItem(0)
    mail.To = email
    
    #assunto
    mail.Subject = 'E-mail automático'
    
    #corpo
    mail.Body = '''
    Prezado {}, 
    Por gentileza validar se deu certo.
    Me manda uma mensagem via WhatsApp para tal confirmação.
    
    Atenciosamente,
    EU
    '''.format(nomes)
    
    #enviar
    mail.Send()
