import datetime
import re
import openpyxl
import pymysql

connectiondb = pymysql.connect(
    host='10.199.102.163',
    user='root',
    passwd='eletro123',
    database='BancoTeste'
)

# Trata texto de data
data = datetime.datetime.now()
data_formatada = data.strftime('%Y%m%d')
ano = data.strftime('%Y')
mes = data.strftime('%m')
data_anterior = data - datetime.timedelta(days=1)
text_data_anterior = data_anterior.strftime('%Y%m%d')
text_data_bd = data_anterior.strftime('%Y-%m-%d %H:%M:%S')
dia = data_anterior.strftime('%d')

# Destino de Origem e Saida
# nome_arquivo = 'R:/logs/SERVIDOR_CCP_NACALA/'+ano+'/'+mes+'/'+dia+'/'+'Console-'+text_data_anterior+'.log'
nome_arquivo = 'Console-20230605.log'
nome_arquivo_excel = 'transicao.xlsx'

def contar_eventos_transicao(texto):
    eventos_transicao = {}
    regex_evento = re.compile(
        r"(Maquina de Chave (\w+) em Transicao)|(Indicação recebida: Maquina de Chave (\w+) em (Transicao))")

    linhas = texto.split('\n')
    for linha in linhas:
        matches = regex_evento.findall(linha)
        for match in matches:
            maquina_chave = match[1] or match[3]
            if maquina_chave in eventos_transicao:
                eventos_transicao[maquina_chave] += 1
            else:
                eventos_transicao[maquina_chave] = 1

    return eventos_transicao


# Ler os dados de um arquivo
try:
    with open(nome_arquivo, 'r', encoding='latin-1') as file:
        texto_dados = file.read()
    eventos_transicao = contar_eventos_transicao(texto_dados)

    # Criar um novo arquivo Excel e adicionar os dados
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Eventos Transicao'

    # Escrever os cabeçalhos
    sheet['A1'] = 'chave'
    sheet['B1'] = 'qtd'
    sheet['C1'] = 'data'

    # Escrever os dados
    row = 2
    for maquina_chave, total_eventos in eventos_transicao.items():
        sheet.cell(row=row, column=1).value = maquina_chave
        sheet.cell(row=row, column=2).value = total_eventos
        sheet.cell(row=row, column=3).value = text_data_bd
        row += 1

        inputdb = connectiondb.cursor()
        SQL_command = 'INSERT INTO TesteTable (nome,cotacao,unidade) VALUES (%s,%s,%s)'
        data = (str(maquina_chave), str(total_eventos), str('usd'))
        inputdb.execute(SQL_command, data)
        connectiondb.commit()

    # Salvar a planilha Excel
    workbook.save(nome_arquivo_excel)

except Exception as e:
    print("Ocorreu um erro:", str(e))


