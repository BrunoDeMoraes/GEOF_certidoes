from pdf2image import convert_from_path
from PIL import Image
import openpyxl
import os
import pytesseract
import re
import time
import datetime

class Certidao:
    def __init__(self, dia, mes, ano):
        self.dia = dia
        self.mes = mes
        self.ano = ano
        self.wb = openpyxl.load_workbook('//hrg-74977/GEOF/CERTIDÕES/Análise/atual.xlsx')
        self.pag = self.wb['PAGAMENTO']
        self.listareferencia = []
        self.referencia = 0
        self.datapag = 'CERTIDÕES PARA {}/{}/{}'.format(self.dia, self.mes, self.ano)
        self.empresas = []
        self.pdf_dir = '//hrg-74977/GEOF/CERTIDÕES/Certidões2/'

    def mensagem_log(self, mensagem):
        with open('//hrg-74977/GEOF/CERTIDÕES/Logs de conferência/{}-{}-{}.txt'.format(self.dia, self.mes, self.ano),
                  'a') as log:
            momento = datetime.datetime.now()
            log.write(f"{mensagem} - {momento.strftime('%d/%m/%Y %H:%M:%S')}\n")

    def mensagem_log_sem_horario(self, mensagem):
        with open('//hrg-74977/GEOF/CERTIDÕES/Logs de conferência/{}-{}-{}.txt'.format(self.dia, self.mes, self.ano),
                  'a') as log:
            log.write(f"{mensagem}\n")

    def pega_referencia(self):
        for linha in self.pag['A1':'P1000']:
            for celula in linha:
                if celula.value != self.datapag:
                    continue
                elif celula.value == self.datapag and celula.coordinate not in self.listareferencia:
                    self.listareferencia.append(celula.coordinate)
        return self.listareferencia

    def pega_fornecedores(self, referencia):
        desloca = 2
        coluna = referencia[0]
        linha = int(referencia[1:])
        while self.pag[coluna + str(linha + desloca)].value != None:
            empresa = self.pag[coluna + str(linha + desloca)].value.split()
            if len(empresa) > 2:
                self.empresas.append(' '.join(empresa[0:len(empresa) - 1]))
            else:
                self.empresas.append(empresa[0])
            desloca += 1
        return self.empresas

    def cria_diretorio(self, fornecedores):
        novos_dir = []
        for emp in fornecedores:
            if os.path.isdir(self.pdf_dir + str(emp)):
                continue
            else:
                os.makedirs(self.pdf_dir + str(emp) + '/Vencidas')
                os.makedirs(self.pdf_dir + str(emp) + '/Imagens')
                novos_dir.append(emp)
        self.mensagem_log(f'\nNúmero de novas pastas criadas: {len(novos_dir)} - {novos_dir}.')
        print(f'Número de novas pastas criadas: {len(novos_dir)} - {novos_dir}.\n')


    def certidoes_n_encontradas(self, fornecedores, orgaos):
        total_faltando = 0
        for emp in fornecedores:
            itens = []
            faltando = []
            os.chdir(self.pdf_dir + str(emp))
            for item in os.listdir(self.pdf_dir + str(emp)):
                itens.append(item.split()[0])
            for orgao in orgaos:
                if orgao not in itens:
                    faltando.append(orgao)
            if faltando != []:
                print(f'Para a empresa {emp} não foram encontradas as certidões {faltando}')
                self.mensagem_log(f'Para a empresa {emp} não foram encontradas as certidões {faltando}')
                total_faltando += 1
        if total_faltando != 0:
            self.mensagem_log(f'Adicione as certidões às respectivas pastas informadas e execute novamente o programa.')
            raise Exception(f'Adicione as certidões às respectivas pastas informadas e execute novamente o programa.')

    def pdf_para_jpg(self, fornecedores, orgaos):
        for emp in fornecedores:
            os.chdir(self.pdf_dir + str(emp))
            for pdf_file in os.listdir(self.pdf_dir + str(emp)):
                if pdf_file.endswith(".pdf") and pdf_file.split()[0] in orgaos:
                    pages = convert_from_path(pdf_file, 300, last_page = 1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")
        self.mensagem_log('\nImagens criadas com sucesso')

    def apaga_imagem(self, fornercedores):
        for emp in fornercedores:
            os.chdir(self.pdf_dir + str(emp))
            for arquivo in os.listdir(self.pdf_dir + str(emp)):
                if arquivo.endswith(".jpg"):
                    os.unlink(self.pdf_dir + str(emp) + f'/{arquivo}')

class Uniao(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(self.pdf_dir + str(emp))
        for imagem in os.listdir(self.pdf_dir + str(emp)):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'UNIÃO':
                certidao = pytesseract.image_to_string(
                    Image.open(f'//hrg-74977/GEOF/CERTIDÕES/Certidões2/{emp}/{imagem}'),
                    lang='por')
                return certidao

    def confere_data(self, certidao):
        texto = []
        padrao = re.compile('do dia (\d\d)/(\d\d)/(\d\d\d\d)')
        emissao_string = padrao.search(certidao)
        texto.append(emissao_string.group().split()[2])
        padrao = re.compile('Válida até (\d\d)/(\d\d)/(\d\d\d\d)')
        vencimento_string = padrao.search(certidao)
        texto.append(vencimento_string.group().split()[2])
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        self.mensagem_log(f'União - emissão {emissao}; válida até: {vencimento}')
        return data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento

class Tst(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(self.pdf_dir + str(emp))
        for imagem in os.listdir(self.pdf_dir + str(emp)):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'TST':
                certidao = pytesseract.image_to_string(
                    Image.open(f'//hrg-74977/GEOF/CERTIDÕES/Certidões2/{emp}/{imagem}'),
                    lang='por')
                return certidao

    def confere_data(self, certidao):
        texto = []
        padrao = re.compile('Expedição: (\d\d)/(\d\d)/(\d\d\d\d)')
        emissao_string = padrao.search(certidao)
        texto.append(emissao_string.group().split()[1])
        padrao = re.compile('Validade: (\d\d)/(\d\d)/(\d\d\d\d)')
        vencimento_string = padrao.search(certidao)
        texto.append(vencimento_string.group().split()[1])
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        self.mensagem_log(f'TST - emissão {emissao}; válida até: {vencimento}')
        return data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento

class Fgts(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(self.pdf_dir + str(emp))
        for imagem in os.listdir(self.pdf_dir + str(emp)):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'FGTS':
                certidao = pytesseract.image_to_string(
                    Image.open(f'//hrg-74977/GEOF/CERTIDÕES/Certidões2/{emp}/{imagem}'),
                    lang='por')
                return certidao

    def confere_data(self, certidao):
        texto = []
        padrao = re.compile('(\d\d)/(\d\d)/(\d\d\d\d) a (\d\d)/(\d\d)/(\d\d\d\d)')
        emissao_string = padrao.search(certidao)
        texto.append(emissao_string.group().split()[0])
        vencimento_string = padrao.search(certidao)
        texto.append(vencimento_string.group().split()[2])
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        self.mensagem_log(f'FGTS - emissão {emissao}; válida até: {vencimento}')
        return data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento

class Gdf(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(self.pdf_dir + str(emp))
        for imagem in os.listdir(self.pdf_dir + str(emp)):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'GDF':
                certidao = pytesseract.image_to_string(
                    Image.open(f'//hrg-74977/GEOF/CERTIDÕES/Certidões2/{emp}/{imagem}'),
                    lang='por')
                return certidao

    def confere_data(self, certidao):
        texto = []
        meses = {'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04', 'Maio': '05', 'Junho': '06',
                 'Julho': '07', 'Agosto': '08', 'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'}
        padrao = re.compile('Brasília, (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                            '(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)')
        emissao_string = padrao.search(certidao)
        datasplit = [emissao_string.group().split()[1], meses[emissao_string.group().split()[3]],
                     emissao_string.group().split()[5]]
        texto.append('/'.join(datasplit))
        padrao = re.compile('Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                            '(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)')
        vencimento_string = padrao.search(certidao)
        datasplit2 = [vencimento_string.group().split()[2], meses[vencimento_string.group().split()[4]],
                     vencimento_string.group().split()[6]]
        texto.append('/'.join(datasplit2))
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        self.mensagem_log(f'GDF - emissão {emissao}; válida até: {vencimento}')
        return data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento