import pandas as pd
import win32com.client as win32

print ("Carregando dados...")
df13 = pd.read_csv("/.txt", sep=';', header=None, names = ['Nome', 'Idade', 'estado civil', 'cidade', 'UF'])
df14 = pd.read_excel("/.xlsx")
df15 = pd.read_excel("/.xlsx")
print ("Dados carregados.")

print ("Executando o programa...")
df13.sort_values(by="Idade")
df14[["Nome","Produto"]]
df15.sort_values(by=["Qtd. Vendida","PrecoUnitario"], ascending=[False, True])
print(df13.shape, df14.shape, df15.shape)

writer = pd.ExcelWriter('planilha_integrada_1.xlsx', engine='xlsxwriter')

planilha1 = pd.DataFrame(df13)
planilha2 = pd.DataFrame(df14)
planilha3 = pd.DataFrame(df15)

planilha1.to_excel(writer, sheet_name='tabela 1')
planilha2.to_excel(writer, sheet_name='tabela 2')
planilha3.to_excel(writer, sheet_name='tabela 3')

writer.save()
print ("Concluído.")
print ("planilha salva no diretório.")

print ("enviando e-mail...")
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To = "email@outlook.com; email@gmail.com"
email.Subject = "Planilha atualizada"
email.HTMLBody = f"""
<p>Senhores, segue arquivo com dados atualizados. </p>
<p>Qualquer dúvida estou a disposição.</p>

<p>Atenciosamente,</p>
<p>autor</p>
"""

anexo = "/.xlsx"
email.Attachments.Add(anexo)

email.Send()
print("Email Enviado.")
