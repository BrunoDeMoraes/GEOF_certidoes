from pdf2image import convert_from_path
from PIL import Image
import openpyxl
import os
import pytesseract
import re
import time
import datetime
import shutil

class Certidao:
    def __init__(self, dia, mes, ano):
        self.dia = dia
        self.mes = mes
        self.ano = ano
        self.wb = openpyxl.load_workbook('//hrg-74977/GEOF/CERTIDÕES/Análise/atual.xlsx')
        self.pag = self.wb['PAGAMENTO']
        self.forn = self.wb['FORNECEDORES']
        self.listareferencia = []
        self.referencia = 0
        self.datapag = f'CERTIDÕES PARA {self.dia}/{self.mes}/{self.ano}'
        self.empresas = []
        self.pdf_dir = '//hrg-74977/GEOF/CERTIDÕES/Certidões2'
        self.percentual = 0

    def mensagem_log(self, mensagem):
        with open(f'//hrg-74977/GEOF/CERTIDÕES/Logs de conferência/{self.ano}-{self.mes}-{self.dia}.txt',
                  'a') as log:
            momento = datetime.datetime.now()
            log.write(f"{mensagem} - {momento.strftime('%d/%m/%Y %H:%M:%S')}\n")

    def mensagem_log_sem_data(self, mensagem):
        with open(f'//hrg-74977/GEOF/CERTIDÕES/Logs de conferência/{self.ano}-{self.mes}-{self.dia}.txt',
                  'a') as log:
            momento = datetime.datetime.now()
            log.write(f"{mensagem} - {momento.strftime('%H:%M:%S')}\n")

    def mensagem_log_sem_horario(self, mensagem):
        with open(f'//hrg-74977/GEOF/CERTIDÕES/Logs de conferência/{self.ano}-{self.mes}-{self.dia}.txt',
                  'a') as log:
            log.write(f"{mensagem}\n")

    def pega_referencia(self):
        for linha in self.pag['A1':'F500']:
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

    def pega_cnpj(self, empresas_a_atualizar):
        for emp in empresas_a_atualizar:
            for linha in self.forn['A6':'A500']:
                for celula in linha:
                    if celula.value == None:
                        continue
                    else:
                        empresa = celula.value.split()
                        nome_da_empresa = ''
                        if len(empresa) > 2:
                            nome_da_empresa = ' '.join(empresa[0:len(empresa) - 1])
                        else:
                            nome_da_empresa = empresa[0]
                        if nome_da_empresa != emp:
                            continue
                        else:
                            cnpj_formatado = str(self.forn['F' + str(celula.row)].value)
                            cnpj_tratado = ''
                            for digito in cnpj_formatado:
                                if digito in '0123456789':
                                    cnpj_tratado += digito
                            empresas_a_atualizar[emp].append(cnpj_tratado)

    def cria_diretorio(self, fornecedores):
        novos_dir = []
        for emp in fornecedores:
            if os.path.isdir(f'{self.pdf_dir}/{str(emp)}'):
                continue
            else:
                os.makedirs(f'{self.pdf_dir}/{str(emp)}/Vencidas')
                os.makedirs(f'{self.pdf_dir}/{str(emp)}/Imagens')
                novos_dir.append(emp)
        self.mensagem_log(f'\nNúmero de novas pastas criadas: {len(novos_dir)} - {novos_dir}.')
        print(f'\nNúmero de novas pastas criadas: {len(novos_dir)} - {novos_dir}.\n')

    def certidoes_n_encontradas(self, fornecedores, orgaos):
        total_faltando = 0
        for emp in fornecedores:
            itens = []
            faltando = []
            os.chdir(f'{self.pdf_dir}/{str(emp)}')
            for item in os.listdir(f'{self.pdf_dir}/{str(emp)}'):
                itens.append(item.split()[0])
            for orgao in orgaos:
                if orgao not in itens:
                    faltando.append(orgao)
            if faltando != []:
                print(f'Para a empresa {emp} não foram encontradas as certidões {faltando}\n')
                self.mensagem_log(f'Para a empresa {emp} não foram encontradas as certidões {faltando}')
                total_faltando += 1
        if total_faltando != 0:
            self.mensagem_log(f'Adicione as certidões às respectivas pastas informadas e execute novamente o programa.')
            raise Exception(f'Adicione as certidões às respectivas pastas informadas e execute novamente o programa.')

    def pdf_para_jpg(self, fornecedores, orgaos):
        for emp in fornecedores:
            os.chdir(f'{self.pdf_dir}/{str(emp)}')
            for pdf_file in os.listdir(f'{self.pdf_dir}/{str(emp)}'):
                if pdf_file.endswith(".pdf") and pdf_file.split()[0] in orgaos:
                    pages = convert_from_path(pdf_file, 300, last_page = 1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")


    def pdf_para_jpg_renomear(self, fornecedores):
        print('CRIANDO IMAGENS:\n')
        for emp in fornecedores:
            os.chdir(f'{self.pdf_dir}/{str(emp)}')
            for pdf_file in os.listdir(f'{self.pdf_dir}/{str(emp)}'):
                if pdf_file.endswith(".pdf"):
                    pages = convert_from_path(pdf_file, 300, last_page=1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")
                    self.percentual += (25 / len(fornecedores))
                    print(f'Total de imagens criadas: {self.percentual}%')
        self.mensagem_log('\nIMAGENS CRIADAS COM SUCESSO!')
        self.percentual = 0


    def gera_nome(self, fornecedores):
        print('\nRENOMEANDO CERTIDÕES:\n\n')
        for emp in fornecedores:
            os.chdir(f'{self.pdf_dir}/{(emp)}')
            origem = f'{self.pdf_dir}/{emp}'
            for imagem in os.listdir(origem):
                if imagem.endswith(".jpg"):
                    certidao = pytesseract.image_to_string(Image.open(f'{origem}/{imagem}'), lang='por')
                    padroes = ['FGTS - CRF', 'Brasília,', 'JUSTIÇA DO TRABALHO', 'MINISTÉRIO DA FAZENDA', 'GOVERNO DO DISTRITO FEDERAL']
                    valores = {'FGTS - CRF': 'FGTS', 'Brasília,': 'GDF', 'JUSTIÇA DO TRABALHO': 'TST',
                               'MINISTÉRIO DA FAZENDA': 'UNIÃO', 'GOVERNO DO DISTRITO FEDERAL':'GDF'}
                    datas = {'FGTS - CRF': 'a (\d\d)/(\d\d)/(\d\d\d\d)',
                             'Brasília,': 'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                             '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)',
                             'JUSTIÇA DO TRABALHO': 'Validade: (\d\d)/(\d\d)/(\d\d\d\d)',
                             'MINISTÉRIO DA FAZENDA': 'Válida até (\d\d)/(\d\d)/(\d\d\d\d)',
                             'GOVERNO DO DISTRITO FEDERAL': 'Válida até (\d\d) de (janeiro)?(fevereiro)?(março)?(Abril)?(maio)?(junho)?'
                                             '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)'}
                    for frase in padroes:
                        if frase in certidao:
                            self.percentual += (25 / len(fornecedores))
                            print(f'{emp} - certidão {valores[frase]} renomeada - Total executado: {self.percentual}%\n')
                            data = re.compile(datas[frase])
                            procura = data.search(certidao)
                            datanome = procura.group()
                            separa = datanome.split('/')
                            junta = '-'.join(separa)
                            if ':' in junta:
                                retira = junta.split(':')
                                volta = ' '.join(retira)
                                junta = volta
                            shutil.move(f'{origem}/{imagem[0:-4]}.pdf', f'{valores[frase]} - {junta}.pdf')
        print('\nPROCESSO DE RENOMEAÇÃO DE CERTIDÕES EXECUTADO COM SUCESSO!')

    def apaga_imagem(self, fornercedores):
        for emp in fornercedores:
            os.chdir(f'{self.pdf_dir}/{str(emp)}')
            for arquivo in os.listdir(f'{self.pdf_dir}/{str(emp)}'):
                if arquivo.endswith(".jpg"):
                    os.unlink(f'{self.pdf_dir}/{str(emp)}/{arquivo}')

class Uniao(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(f'{self.pdf_dir}/{str(emp)}')
        for imagem in os.listdir(f'{self.pdf_dir}/{str(emp)}'):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'UNIÃO':
                certidao = pytesseract.image_to_string(
                    Image.open(f'{self.pdf_dir}/{emp}/{imagem}'),
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
        self.mensagem_log_sem_data(f'UNIÃO - emissão: {emissao}; válida até: {vencimento}')
        print(f'    UNIÃO - emissão: {emissao}; válida até: {vencimento}')
        return data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento

class Tst(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(f'{self.pdf_dir}/{str(emp)}')
        for imagem in os.listdir(f'{self.pdf_dir}/{str(emp)}'):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'TST':
                certidao = pytesseract.image_to_string(
                    Image.open(f'{self.pdf_dir}/{emp}/{imagem}'),
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
        self.mensagem_log_sem_data(f'TST   - emissão: {emissao}; válida até: {vencimento}')
        print((f'    TST   - emissão: {emissao}; válida até: {vencimento}'))
        return data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento

class Fgts(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(f'{self.pdf_dir}/{str(emp)}')
        for imagem in os.listdir(f'{self.pdf_dir}/{str(emp)}'):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'FGTS':
                certidao = pytesseract.image_to_string(
                    Image.open(f'{self.pdf_dir}/{emp}/{imagem}'),
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
        self.mensagem_log_sem_data(f'FGTS  - emissão: {emissao}; válida até: {vencimento}')
        print(f'    FGTS  - emissão: {emissao}; válida até: {vencimento}')
        return data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento

class Gdf(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(f'{self.pdf_dir}/{str(emp)}')
        for imagem in os.listdir(f'{self.pdf_dir}/{str(emp)}'):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'GDF':
                certidao = pytesseract.image_to_string(
                    Image.open(f'{self.pdf_dir}/{emp}/{imagem}'),
                    lang='por')
                return certidao

    def confere_data(self, certidao):
        texto = []
        meses = {'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04', 'Maio': '05', 'Junho': '06',
                 'Julho': '07', 'Agosto': '08', 'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'}
        meses2 = {'janeiro': '01', 'fevereiro': '02', 'março': '03', 'abril': '04', 'maio': '05', 'junho': '06',
                 'julho': '07', 'agosto': '08', 'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12'}
        if "GOVERNO" not in certidao:
            padrao = re.compile('Brasília, (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                                '(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)')
            emissao_string = padrao.search(certidao)
            datasplit = [emissao_string.group().split()[1], meses[emissao_string.group().split()[3]],
                         emissao_string.group().split()[5]]
            texto.append('/'.join(datasplit))
            padrao = re.compile(
                'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                '(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)')
            vencimento_string = padrao.search(certidao)
            datasplit2 = [vencimento_string.group().split()[2], meses[vencimento_string.group().split()[4]],
                          vencimento_string.group().split()[6]]
            texto.append('/'.join(datasplit2))
        else:
            padrao = re.compile('Certidão emitida via internet em (\d\d)/(\d\d)/(\d\d\d\d)')
            emissao_string = padrao.search(certidao)
            texto.append(emissao_string.group().split()[5])
            padrao = re.compile('Válida até (\d\d) de (janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?(julho)?(agosto)?'
                                '(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)')
            vencimento_string = padrao.search(certidao)
            datasplit2 = [vencimento_string.group().split()[2], meses2[vencimento_string.group().split()[4]],
                         vencimento_string.group().split()[6]]
            texto.append('/'.join(datasplit2))
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        self.mensagem_log_sem_data(f'GDF   - emissão: {emissao}; válida até: {vencimento}')
        print((f'    GDF   - emissão: {emissao}; válida até: {vencimento}'))
        return data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento
