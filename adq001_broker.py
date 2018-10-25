# encoding: iso-8859-1
# encoding: win-1252
# encoding: utf-8
"""
Este script interpreta os dados contidos no arquivo ADQ001.TXT e salva os dados em CSV e XLSX.
"""
from datetime import datetime as dt
import pandas as pd
class Adq001broker():
    """
    Classe que interpreta os arquivos ADQ001 conforme protocolo especifico.
    :params:
        file_env = arquivo texto contendo os dados
    :vars:
        file: arquivo txt carregado em memoria
        band: bandeiras dos cartoes
        now: data e hora atual
        csv_name: nome que sera dado ao arquivo de saida
    :methods:
        triagem: verifica se a linha se refere ao header, detail ou tail
        header: interpreta os valores de header
        detail: interpreta os valores de detail
        tail: interpreta os valores de tail
        save_csv: salva os dados no arquivo de saida no formato csv
        save_xlsx: salva os dados no arquivo de saida no formato xlsx
    """
    def __init__(self, file_env):
        self.cnt = 0
        self.file_env = file_env
        with open(self.file_env, 'r', encoding='utf-8') as file:
            self.env = file.readlines()
            file.close()
        self.band = {'003': 'Mastercard', '004': 'Visa', '005': 'Diners Club',\
					'006': 'American Express', '008': 'Elo', '009': 'Alelo', \
					'010': 'Cabal', '011': 'Agiplan', '012': 'Aura', '013': 'Banescard', \
					'014': 'Calcard', '015': 'Credsystem', '016': 'Cup', '017': 'Redesplan',\
					'018': 'Sicred', '019': 'Sorocred', '020': 'Verdecard', '021': 'Hipercard',\
					'022': 'Avista', '023': 'Credz', '024': 'Discover', '025': 'Maestro',\
					'026': 'Visa Electron', '027': 'Elo débito', '028': 'Sicredi débito',\
					'029': 'Hiper crédito', '030': 'Cabal débito', '031': 'JCB', '032': 'Ticket',\
					'033': 'Sodexo', '034':'VR', '035': 'Policard', '036': 'Valecard',\
					'037': 'Goodcard', '038': 'Greencard', '039': 'Coopercard',\
					'040': 'Verocheque', '041': 'Nutricash', '042': 'Banricard',\
					'043': 'Banescard débito', '044': 'Sorocred pré-pago',\
					'045': 'Mastercard pré-pago', '046': 'Visa pré-pago', '047': 'Ourocard'}
        self.now = dt.now()
        self.csv_name = "ADQ001_" + (str(self.now.microsecond)) + ".csv"
        self.triagem()
        self.save_xlsx()
    def triagem(self):
        """
        Este metodo verifica qual o tipo de dado contido em cada linha do arquivo TXT e chama o
		metodo correto para interpreta-la.
        :vars:
            linha: string que recebe linha a linha do conteudo do arquivo txt.
            cont: contador das linhas lidas.
        """
        cont = 0
        for linha in self.env:
            if linha[0] == '0':
                self.header(cont, linha)
            elif linha[0] == '1':
                self.detail(cont, linha)
            elif linha[0] == '9':
                self.tail(cont, linha)
            else:
                print("Unknown value at line %i" % (cont  +  1))
            cont += 1
        print(cont)
    def header(self, cont, linha):
        """
        Este metodo interpreta os dados de cabecalho do arquivo txt.
        :params:
            cont: contador proveniente da triagem.
            linha: string a ser interpreta proveniente da triagem.
        :vars:
            point: indice do posicionamento dentro das strings para precorrer os dados
            conforme o protocolo.
            cnt: contador.
            cabecalho: Titulos das colunas.
            filler: valores em branco no final de cada linha
            as demais variaveis sao referentes aos valores de cada coluna, de acordo com o
            protocolo.
        """
        point = 0
        self.cnt = cont  +  1
        point += 1
        id_emissor = linha[point: point  +  30]
        print("Identificacao do emissor: %s" % id_emissor)
        point += 30
        id_destinatario = linha[point: point  +  30]
        print("Identificacao do destinatário: %s" % id_destinatario)
        point += 30
        cod_parceiro = linha[point: point  +  2]
        print("Código do parceiro: %s" % cod_parceiro)
        point += 2
        file_dt = linha[point: point  +  14]
        print("Data Hora: %s-%s-%s %s:%s:%s" % (file_dt[0:4], file_dt[4:6], file_dt[6:8], \
		        file_dt[8:10], file_dt[10:12], file_dt[12:]))
        point += 14
        tipo_arranjo = linha[point: point  +  1]
        print("Tipo de operação para o arranjo que está sendo liquidado para o cliente: %s" \
		        %tipo_arranjo)
        point += 1
        cod_validacao = linha[point: point  +  200]
        print("Código de validação: %s" % cod_validacao)
        point: point  +  200
        filler = linha[point: ]
        print("Filler: %i" % len(filler))
        cabecalho = ("Identificador,Data PG, CPF/CNPJ, Nome, Tipo Cliente,Valor PG, "
                     "ID Instrução PG, Tipo Conta, Número Banco, Agência, Conta, Conta Pagamento, "
                     "bandeira, Filler")
        self.save_csv(cabecalho)
    def detail(self, cont, linha):
        """
        Este metodo interpreta os dados de detalhes do arquivo txt.
        :params:
            cont: contador proveniente da triagem.
            linha: string a ser interpreta proveniente da triagem.
        :vars:
            v: insere a virgula para criacao do csv.
            point: indice do posicionamento dentro das strings para precorrer os dados conforme
            o protocolo.
            cnt: contador
            identificador_linha: identifica o tipo de operacao, conforme o protocolo
            filler: valores em branco no final de cada linha
            as demais variaveis sao referentes aos valores de cada coluna, de acordo com o
            protocolo.
        """
        comma = ','
        point = 0
        self.cnt = cont  +  1
        identificador_linha = linha[point: point  +  1]
        point += 1
        data_pagamento = linha[point: point  +  4]  +  "-"  +  linha[point  +  4: point  +  6]  \
		                    +  "-" +  linha[point  +  6:point  +  8]
        point += 8
        doc_cliente = linha[point: point  +  14]
        point += 14
        nome_cliente = linha[point: point  +  50]
        point += 50
        tipo_cliente = linha[point: point  +  1]
        point += 1
        valor_pagam = "%.2f" %(int(linha[point: point  +  19]) / 100)
        point += 19
        id_instrucao_pagam = linha[point: point  +  18]
        point += 18
        tipo_conta_cliente = linha[point: point  +  2]
        point += 2
        banco_cliente = linha[point: point  +  4]
        point += 4
        agencia_cliente = linha[point: point  +  4]
        point += 4
        conta_cliente = linha[point: point  +  13]
        point += 13
        num_conta_pg_cliente = linha[point: point  +  20]
        point += 20
        bandeira = self.band[linha[point: point  +  3]]
        point += 3
        filler = str(len(linha[point:]))

        self.save_csv(identificador_linha + comma + data_pagamento + comma + doc_cliente + comma + \
                     nome_cliente + comma + tipo_cliente + comma + valor_pagam + comma + \
                     id_instrucao_pagam + comma + tipo_conta_cliente + comma + banco_cliente\
                     + comma + agencia_cliente + comma + conta_cliente + comma + \
					 num_conta_pg_cliente + comma + bandeira + comma + filler)
    def tail(self, cont, linha):
        """
        Este metodo interpreta os dados de rodape do arquivo txt.
        :params:
            cont: contador proveniente da triagem.
            linha: string a ser interpreta proveniente da triagem.
        :vars:
            v: insere a virgula para criacao do csv.
            point: indice do posicionamento dentro das strings para precorrer os dados conforme
            o protocolo.
            cnt: contador.
            filler: valores em branco no final de cada linha as demais variaveis sao referentes
            aos valores de cada coluna, de acordo com o protocolo.
        """
        point = 0
        self.cnt = cont + 1
        point += 1
        qtd_lancamentos = linha[point: point + 6]
        print("Quantidade de lançamentos: %s" %qtd_lancamentos)
        point += 6
        soma_valores = linha[point: point + 19]
        print("Somatório dos valores das operações: R$%s" %str(float(soma_valores)/100))
        point += 19
        filler = linha[point:]
        print("Filler: %i" %len(filler))
    def save_csv(self, linha):
        """
        Este metodo cria o arquivo csv de saida.
        :params:
            linha: linha que sera gravada no arquivo de saida.
        :vars:
            csvfile: objeto arquivo que sera usado para gravar o csv.
        """
        with open(self.csv_name, 'a') as csvfile:
            csvfile.write(linha)
            csvfile.write('\n')
    def save_xlsx(self):
        """
        Este metodo cria um arquivo xlsx de saida utilizando o Pandas.
        :vars:
            csv: arquivo csv gerado a partir da interpretacao do arquivo txt.
            writer: escreve o arquivo xlsx (xlsxwriter é o engine escolhido para a gravacao
            dos dados no arquivo de saida).
            work: cria o workbook para manipulacao dos dados em xlsx.
            wsheet: seleciona a pagina(planilha, aba ou sheet) onde os dados serao gravados
            os dados.
            format(x): estipula o formato dos dados gravados em cada coluna.
        """
        csv = pd.read_csv(self.csv_name, encoding='cp1252')
        writer = pd.ExcelWriter(self.csv_name + ".xlsx", engine='xlsxwriter')
        csv.to_excel(writer, sheet_name='Sheet1', index=None)
        work = writer.book
        wsheet = writer.sheets['Sheet1']
        format_1 = work.add_format({'num_format':'####00000000000'})
        format_2 = work.add_format({'num_format':'0.00'})
        format_3 = work.add_format({'num_format':'dd/mm/yyyy'})
        format_a = work.add_format({'bold':True, 'font_color':'green'})
        wsheet.set_column('A:A', 12, None)
        wsheet.set_column('B:B', 10, format_3)
        wsheet.set_column('C:C', 15, format_1)
        wsheet.set_column('D:D', 50, format_a)
        wsheet.set_column('E:E', 11.43, format_a)
        wsheet.set_column('F:F', 8, format_2)
        wsheet.set_column('G:G', 16, format_1)
        wsheet.set_column('I:I', 13.43, format_a)
        wsheet.set_column('M:M', 11, format_a)
