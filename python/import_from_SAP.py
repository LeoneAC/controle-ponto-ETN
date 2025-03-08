# %%
### Script para pegar as marcações de ponto do SAP e adicionar no Excel de controle de ponto
import pandas as pd
import pdfplumber
import PyPDF2
import numpy as np
import xlwings
from datetime import datetime
# import time
from re import search as re_search
import locale
import os
import win32com.client

# Para gerar o .exe: pyinstaller --onefile --name "import_from_SAP" import_from_SAP.py

# %%
#"Definição" de constantes
VERSAO = 'v1.0.0-python'
ARQUIVO_PDF_DEFAULT = 'smart.pdf'
ARQUIVO_EXCEL_DEFAULT_INI = 'Controle de Horas '
ARQUIVO_EXCEL_DEFAULT_FIM = '.xlsm'

# %%
#Função para verificar se o Excel já está aberto manualmente e se o arquivo está na lista
def is_workbook_open(filename):
    # excel_instances = win32com.client.GetObject(None, "Excel.Application")
    excel_instances = win32com.client.Dispatch("Excel.Application")
    for workbook in excel_instances.Workbooks:
        if workbook.FullName.lower() == os.path.abspath(filename).lower():
            if workbook.ReadOnly: # Essa parte vai crashar o arquivo, mas vai conseguir pelo menos finalizar o import, o ideal é que o arquivo já tenha sido salvo habilitado para edição certinho pelo menos uma vez
                excel_instances.EnableEvents = True
            return True
    return False

# %%
#Função para pegar retornar a posição da linha extra a ser adicionada no fim da tabela
def extra_line(page):
    text_data = page.extract_words()
    if text_data:
        return text_data[-1]["bottom"]
    return 0

# %%
print(f'## Versão: {VERSAO} ##\n')
print('Este programa serve para importar os dados exportados pelo SAP para o Excel de Controle de Ponto.\n')
print('!!!ATENÇÃO!!!\n'\
    'Para que o script funcione, é importante que o arquivo Excel, o PDF e o próprio script estejam todos numa mesma pasta.')

# %%
'''Pegando o nome do arquivo .pdf e extrai algumas informações'''
while True:
    arquivo_pdf = input('\nDigite o nome do PDF com as marcações de ponto:\n'\
                    f'(O arquivo padrão é "{ARQUIVO_PDF_DEFAULT}", para usá-lo, apenas tecle Enter)\n>')\
                    or ARQUIVO_PDF_DEFAULT
    if arquivo_pdf[-4:] != '.pdf': arquivo_pdf = arquivo_pdf + '.pdf'
    try:
        with open(arquivo_pdf, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file) # Lê o pdf
            npages = len(pdf_reader.pages) # Número de páginas
            page1 = pdf_reader.pages[0].extract_text() # Usa a primeira página porque sei que a info está lá
        break
    except FileNotFoundError:
        print(f'\nO arquivo "{arquivo_pdf}" não foi encontrado.')

# %%
'''Extrai as informações de ano e mês do pdf'''
# ini = page1.find('\n')
# fim = page1.find('\n', ini+1)
# page1 = page1[(ini+1):fim]
match = re_search(r'\d{2}\.\d{2}\.\d{4}',page1)
data = datetime.strptime(match.group(),'%d.%m.%Y')
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
mes = datetime.strftime(data,'%b').capitalize()
ano = datetime.strftime(data,'%Y')

# %%
'''Pegando o nome do arquivo excel'''
while True:
    arquivo_excel = input('\nDigite o nome da sua planilha do EXCEL de controle de pontos:\n'\
                    f'(Se o nome da sua planilha for "{ARQUIVO_EXCEL_DEFAULT_INI}{ano}{ARQUIVO_EXCEL_DEFAULT_FIM}", apenas tecle Enter)\n>')\
                    or ARQUIVO_EXCEL_DEFAULT_INI + ano + ARQUIVO_EXCEL_DEFAULT_FIM
    if arquivo_excel[-5:] != '.xlsm' and arquivo_excel[-5:] != '.xlsx':
        arquivo_excel = arquivo_excel + '.xlsm'
    if os.path.exists(arquivo_excel):
        #Arquivo excel existe
        break
    else:
        print(f'\nO arquivo "{arquivo_excel}" não foi encontrado.')

# %%
print('\nProcessando...')

# %%
'''Extrai as tabelas do pdf (pega a principal tabela de cada página, que já o que nos atende)'''
DF = [] # array com os dataframes
with pdfplumber.open(arquivo_pdf) as pdf:
    # Selecione a página desejada
    for p in range(npages):
        page = pdf.pages[p]
        # Extraia a tabela da página (adição de linha abaixo para corrigir a falta de linha do SAP no das tabelas com meses incompletos)
        table = page.extract_table(table_settings={"explicit_horizontal_lines": [extra_line(page)]})
        if table:
            # Converta a tabela para DataFrame
            df = pd.DataFrame(table[1:], columns=table[0])
            DF.append(df)

# %%
# Seção de testes, desconsidere
'''
with pdfplumber.open(arquivo_pdf) as pdf:           
    page = pdf.pages[0]
    im = page.to_image()
    table = page.extract_table(table_settings={"explicit_horizontal_lines": [781.206]}) # antigo: 775
    if table:
        df = pd.DataFrame(table[1:], columns=table[0])

horizontal_lines = [line for line in page.lines if line["y1"] == line["y0"]]
last_hl = horizontal_lines[-1]
for h in horizontal_lines:
    if h["y1"] > last_hl["y1"]:
        last_hl = h
display(last_hl)
display(last_hl["y1"])
text_data = page.extract_words()
last_text = text_data[-1]
display(last_text["bottom"])
display(last_text)
display(df)
im.debug_tablefinder(table_settings={"explicit_horizontal_lines": [last_text["bottom"]]})
# im.debug_tablefinder(table_settings={"explicit_horizontal_lines": [last_hl["y1"]+0]})

# display(list(df.columns))
# display(table[0])

# TAB = DF[0]
# display(TAB.columns)
# display(list(range(1,len(DF))))
'''

# %%
'''Junta cada tabela extraída em uma só, eliminando as tabelas que não nos interessa (que não contêm registro de ponto)'''
TAB = DF[0] # Estou definindo a primeira tabela como referência porque eu sei com certeza que ela é a tabela correta de pontos
for i in range(1,len(DF)):
    tab2 = DF[i]
    if list(TAB.columns) == list(tab2.columns):
        TAB = pd.concat([TAB, tab2], ignore_index=True)

# %%
'''Tratamento da tabela maior, que contém todos os registros de pontos, para que só sobre linhas com Marcação de ponto, nada mais'''
#Seleciona as colunas de interesse
tab = TAB.iloc[:,np.r_[0,2:5]].copy()
#Expande a informação de 'Dia' para todas as linhas
for i in range(1,len(tab)):
    if tab.loc[i,'Dia'] == '':
        tab.loc[i,'Dia'] = tab.loc[i-1,'Dia']
#Renomeia as colunas e cria o restante das colunas de ponto (de 2 colunas para 4)
tab.rename(columns={tab.columns[1]: 'Descrição', 'De':'Entrada', 'Até':'Saída Almoço'}, inplace=True)
tab['Retorno Almoço'] = ''
tab['Saída'] = ''
#Filtra apenas as linhas com Marcação de ponto e regenera os índices das linhas
tab = tab[tab['Descrição']=='Marcação']
tab.reset_index(drop = True, inplace=True)
#Trata os dias que possuem 2 marcações e os que possuem 4 para que só haja uma linha por dia
for i in range(len(tab)-1):
    if tab.at[i,'Descrição'] == 'Marcação':
        if tab.at[i,'Dia'] == tab.loc[i+1,'Dia']:
            tab.loc[i,['Retorno Almoço', 'Saída']] = tab.loc[i+1,['Entrada', 'Saída Almoço']].values
            tab.at[i+1,'Descrição'] = 'Apagar'
        else:
            tab.at[i,'Saída'] = tab.at[i,'Saída Almoço']
            tab.at[i,'Saída Almoço'] = ''
#Filtra novamente para retirar as linhas que não são mais necessárias e regenera os índices
tab = tab[tab['Descrição']=='Marcação']
tab.reset_index(drop = True, inplace=True)
#Converte os dias para dado numérico
tab['Dia'] = pd.to_numeric(tab['Dia'])
#Acrescenta os dias que não têm marcação para a tabela ficar com tamanho constante de 31
if tab['Dia'].tail(1).item() != 31: # Marca o último dia para evitar erros
    tab.loc[len(tab)] = ''
    tab.loc[tab.index[-1], 'Dia'] = 31
for i in range(31):
    if tab.at[i,'Dia'] != i+1:
        new_line = {col: '' for col in tab.columns}  # Gera uma linha vazia
        new_line['Dia'] = i+1  # Adicione valor à primeira coluna
        new_line = pd.DataFrame([new_line])
        tab = pd.concat([tab.iloc[:i], new_line, tab.iloc[i:]]).reset_index(drop=True)
#Converte os Pontos em tipo tempo para inserir no excel
time_tab = tab.iloc[:, -4:].apply(lambda x: x.apply(lambda y: datetime.strptime(y.split()[0], '%H:%M') if y else ''))

# %%
'''Seção de visualização'''
'''
pd.set_option('display.max_rows', 136)
# display(TAB)
display(tab)
# display(time_tab)
# display(time_tab.iloc[:, -4:].applymap(lambda x: x.strftime('%H:%M') if x else ''))
pd.reset_option('display.max_rows')
'''

# %%
'''Insere os dados na planilha excel (pacote openpyxl não deu certo, ele transforma as fórmulas em valores)'''
'''
excel_file = 'Controle de Horas 2025 - Com Macro.xlsm'
mes = 'Nov'
workbook = openpyxl.load_workbook(excel_file, keep_vba=True)
sheet = workbook[mes]

# Escrever os valores no range
for i, row in enumerate(time_tab.iloc[:,-4:].itertuples(index=False), start=2): # começa na linha 2
    for j, ponto in enumerate(row, start=4): # começa na coluna D
        if ponto:
            sheet.cell(row=i, column=j, value=(ponto.hour*60+ponto.minute)/(24*60))

workbook.save(excel_file)
'''

# %%
'''Abre o excel'''
# Tenta encontrar o arquivo já aberto
if is_workbook_open(arquivo_excel):
    #Esse trecho abaixo é uma gambiarra, não consigo passar um objeto do win32com.client (workbook_com, antigo retorno da função is_workbook_open)
    #para o xlwings (app), então tive que ativar o workbook_com e capturar a janela ativa com apps.active
    # workbook_com.Activate() # Esse 
    # time.sleep(1) # Achei que pudesse ser necessário, mas parece que deu certo sem
    # app = xlwings.apps.active
    # workbook = [book for book in app.books if book.fullname.lower() == os.path.abspath(arquivo_excel).lower()][0]  # Localiza o workbook no xlwings

    #Resolvi fazer essa abordagem que deu certo e me pareceu mais segura e limpa
    workbook = xlwings.Book(os.path.abspath(arquivo_excel)) # Abre o workbook pelo seu caminho
    workbook.activate() # Ativa a planilha correta
    app = xlwings.apps.active # Captura o Excel ativo
    app.api.WindowState = -4137 # Maximiza a janela do Excel
    fechar = False # O arquivo já estava aberto, então não precisamos fechar ao final
else:
# Se não estiver aberto, cria uma nova instância oculta do Excel
    app = xlwings.App(visible=False) # Cria uma nova instância invisível
    workbook = app.books.open(arquivo_excel) # Abre o arquivo Excel
    fechar = True # Indica que criamos uma nova instância e que precisamos fechar no final do código

# %%
'''Insere os dados na planilha excel'''
sheet = workbook.sheets[mes]  # Seleciona a planilha
sheet.activate() # Mostra a planilha

# time_tab = time_tab.values.tolist() # Converte para lista de listas

# Escrever os valores no range
for r, row in enumerate(time_tab.itertuples(index=False), start=2): # começa na linha 2
    for c, ponto in enumerate(row, start=4): # começa na coluna D
        if ponto:
            sheet.range((r,c)).value = (ponto.hour*60+ponto.minute)/(24*60)
            # time.sleep(0.05) # delay de 0,05s

# %%
'''Fecha o excel'''
# Salva e fecha apenas se criamos uma nova instância
if fechar:
    workbook.save()     # Salva as alterações
    workbook.close()    # Fecha o arquivo
    app.quit()          # Fecha o Excel

# %%
print(f'\nInformações de ponto de "{arquivo_pdf}" carregadas em "{arquivo_excel}" com sucesso!')
input("\nPressione Enter para fechar.")


