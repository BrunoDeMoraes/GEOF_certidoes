from pdf2image import convert_from_path
from PIL import Image
import openpyxl
import os
import pytesseract
import re
import time
import datetime
import shutil
import PyPDF2

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
        self.empresas = {}
        self.pdf_dir = '//hrg-74977/GEOF/CERTIDÕES/Certidões2'
        self.percentual = 0
        self.lista_de_cnpj = {}
        self.orgaos = ['UNIÃO', 'TST', 'FGTS', 'GDF']
        self.empresasdic = {}
        self.empresas_a_atualizar = {}

    def mensagem_log(self, mensagem):
        with open(f'//hrg-74977/GEOF/CERTIDÕES/Logs de conferência/{self.ano}-{self.mes}-{self.dia}.txt',
                  'a') as log:
            momento = datetime.datetime.now()
            log.write(f"{mensagem} - {momento.strftime('%d/%m/%Y %H:%M:%S')}\n")
            print(f"{mensagem} - {momento.strftime('%d/%m/%Y %H:%M:%S')}")

    def mensagem_log_sem_data(self, mensagem):
        with open(f'//hrg-74977/GEOF/CERTIDÕES/Logs de conferência/{self.ano}-{self.mes}-{self.dia}.txt',
                  'a') as log:
            momento = datetime.datetime.now()
            log.write(f"{mensagem} - {momento.strftime('%H:%M:%S')}\n")
            print(f"{mensagem} - {momento.strftime('%H:%M:%S')}")

    def mensagem_log_sem_horario(self, mensagem):
        with open(f'//hrg-74977/GEOF/CERTIDÕES/Logs de conferência/{self.ano}-{self.mes}-{self.dia}.txt',
                  'a') as log:
            log.write(f"{mensagem}\n")
            print(f"{mensagem}")

    def pega_referencia(self):
        pasta_de_trabalho = f'//hrg-74977/GEOF/HRG/PDPAS 2020/PAGAMENTO/{self.ano}-{self.mes}-{self.dia}'
        if os.path.exists(f'{pasta_de_trabalho}'):
            print('Pasta para inclusão de arquivos de pagamento localizada.')
        else:
            os.makedirs(f'{pasta_de_trabalho}')
            print('Pasta para inclusão de arquivos de pagamento criada com sucesso.')
        for linha in self.pag['A1':'F500']:
            for celula in linha:
                if celula.value != self.datapag:
                    continue
                elif celula.value == self.datapag and celula.coordinate not in self.listareferencia:
                    self.listareferencia.append(celula.coordinate)
        return self.listareferencia

    def analisa_referencia(self):
        self.pega_referencia()
        if len(self.listareferencia) == 0:
            self.mensagem_log('\nData específicada não encontrada')
            raise Exception('Data não encontrada!')
        elif len(self.listareferencia) > 1:
            self.mensagem_log('Data informada em multiplicidade')
            print(f'A data especificada foi encontrada nas células {self.listareferencia} da planilha de pagamentos.'
                            f'\nApague os valores duplicados e execute o programa novamente.')
            raise Exception(f'A data especificada foi encontrada nas células {self.listareferencia} da planilha de pagamentos.'
                            f'\nApague os valores duplicados e execute o programa novamente.')
        else:
            self.mensagem_log(f'\nReferência encontrada na célula {self.listareferencia[0]}')

    def pega_fornecedores(self):
        referencia = self.listareferencia[0]
        desloca = 2
        coluna = referencia[0]
        linha = int(referencia[1:])
        while self.pag[coluna + str(linha + desloca)].value != None:
            empresa = self.pag[coluna + str(linha + desloca)].value.split()
            if len(empresa) > 2:
                self.empresas[' '.join(empresa[0:len(empresa) - 1])] = [' '.join(empresa)]
            else:
                self.empresas[empresa[0]] = [' '.join(empresa)]
            desloca += 1
        return self.empresas

    def inclui_cnpj_em_fornecedores(self):
        for emp in self.empresas:
            for linha in self.forn['A6':'A500']:
                for celula in linha:
                    if celula.value != self.empresas[emp][0]:
                        continue
                    else:
                        self.empresas[emp].append(self.forn[f'F{celula.row}'].value)
                        cnpj_formatado = self.empresas[emp][1]
                        cnpj_tratado = ''
                        for digito in cnpj_formatado:
                            if digito in '0123456789':
                                cnpj_tratado += digito
                        self.empresas[emp].append(cnpj_tratado)
        return self.empresas

    def dados_completos_dos_fornecedores(self):
        self.pega_fornecedores()
        self.inclui_cnpj_em_fornecedores()

    def listar_cnpjs(self):
        for emp in self.empresas:
            for linha in self.forn['F6':'F500']:
                for celula in linha:
                    if celula.value == None:
                        continue
                    else:
                        nome_da_empresa = self.forn[f'A{celula.row}'].value.split()
                        self.lista_de_cnpj[celula.value] = ' '.join(nome_da_empresa[0:len(nome_da_empresa) - 1])
        return self.lista_de_cnpj

    def cria_diretorio(self):
        novos_dir = []
        for emp in self.empresas:
            if os.path.isdir(f'{self.pdf_dir}/{str(emp)}'):
                continue
            else:
                os.makedirs(f'{self.pdf_dir}/{str(emp)}/Vencidas')
                os.makedirs(f'{self.pdf_dir}/{str(emp)}/Imagens')
                novos_dir.append(emp)
        self.mensagem_log(f'\nNúmero de novas pastas criadas: {len(novos_dir)} - {novos_dir}.')

    def certidoes_para_pagamento(self):
        pagamento_por_data = f'//hrg-74977/GEOF/CERTIDÕES/Pagamentos/{self.ano}-{self.mes}-{self.dia}'
        if os.path.exists(f'{pagamento_por_data}'):
            print('Já existe pasta contendo certidões para pagamento na data informada.')
        else:
            os.makedirs(pagamento_por_data)
            for emp in self.empresas:
                pasta_da_empresa = f'{self.pdf_dir}/{str(emp)}'
                os.makedirs(f'{pagamento_por_data}/{emp}')
                os.chdir(f'{pasta_da_empresa}')
                for pdf_file in os.listdir(f'{pasta_da_empresa}'):
                    if pdf_file.endswith(".pdf"):
                        shutil.copy(f'{pasta_da_empresa}/{pdf_file}', f'{pagamento_por_data}/{emp}/{pdf_file}')
            self.mensagem_log_sem_horario(f'As certidões referentes ao pagamento com data limite para a data de {self.dia}/{self.mes}/{self.ano} foram transferidas para respectiva pasta de pagamento.')


    def certidoes_n_encontradas(self):
        total_faltando = 0
        for emp in self.empresas:
            itens = []
            faltando = []
            os.chdir(f'{self.pdf_dir}/{str(emp)}')
            for item in os.listdir(f'{self.pdf_dir}/{str(emp)}'):
                itens.append(item.split()[0])
            for orgao in self.orgaos:
                if orgao not in itens:
                    faltando.append(orgao)
            if faltando != []:
                self.mensagem_log(f'Para a empresa {emp} não foram encontradas as certidões {faltando} - CNPJ: {self.empresas[emp][2]}')
                total_faltando += 1
        if total_faltando != 0:
            self.mensagem_log(f'Adicione as certidões às respectivas pastas informadas e execute novamente o programa.')
            raise Exception(f'Adicione as certidões às respectivas pastas informadas e execute novamente o programa.')

    def pdf_para_jpg(self):
        for emp in self.empresas:
            os.chdir(f'{self.pdf_dir}/{str(emp)}')
            for pdf_file in os.listdir(f'{self.pdf_dir}/{str(emp)}'):
                if '00.MERGE' in pdf_file:
                    if not os.path.isdir(f'{self.pdf_dir}/{str(emp)}/Merge'):
                        os.makedirs(f'{self.pdf_dir}/{str(emp)}/Merge')
                        shutil.move(pdf_file, f'{self.pdf_dir}/{str(emp)}/Merge/{pdf_file}')
                    else:
                        shutil.move(pdf_file, f'{self.pdf_dir}/{str(emp)}/Merge/{pdf_file}')
                if pdf_file.endswith(".pdf") and pdf_file.split()[0] in self.orgaos:
                    pages = convert_from_path(pdf_file, 300, last_page = 1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")

    def analisa_certidoes(self):
        objUniao = Uniao(self.dia, self.mes, self.ano)
        objTst = Tst(self.dia, self.mes, self.ano)
        objFgts = Fgts(self.dia, self.mes, self.ano)
        objGdf = Gdf(self.dia, self.mes, self.ano)
        lista_objetos = [objUniao, objTst, objFgts, objGdf]
        self.mensagem_log('\nInicio da conferência de datas de emissão e vencimento:')
        print(f'Total executado: {self.percentual}%')

        for emp in self.empresas:
            empresadic = {}
            index = 0
            self.mensagem_log(f'\n{emp}')
            for objeto in lista_objetos:
                objeto.empresas = self.empresas
                objeto.lista_de_cnpj = self.lista_de_cnpj
                certidao = objeto.pega_string(emp)
                self.percentual += (25 / len(self.empresas))
                print(f'   Total executado: {self.percentual}%')
                val, cnpj_para_comparação = objeto.confere_data(certidao)
                if val == True and cnpj_para_comparação == self.empresas[emp][1]:
                    empresadic[self.orgaos[index]] = 'OK'
                elif cnpj_para_comparação != self.empresas[emp][1]:
                    empresadic[self.orgaos[index]] = 'CNPJ-ERRO'
                else:
                    empresadic[self.orgaos[index]] = 'INCOMPATÍVEL'
                index += 1
            self.empresasdic[emp] = empresadic

    def atualizar(self):
        numerador = 0
        for emp in self.empresasdic:
            self.mensagem_log_sem_horario(f'{numerador + 1 :>2} - {emp}\n{self.empresasdic[emp]}\n')
            numerador += 1
        for emp in self.empresasdic:
            certidoes_a_atualizar = []
            for orgao in self.empresasdic[emp]:
                if self.empresasdic[emp][orgao] == 'INCOMPATÍVEL' or self.empresasdic[emp][orgao] == 'CNPJ-ERRO':
                    certidoes_a_atualizar.append(orgao)
            if len(certidoes_a_atualizar) > 0:
                self.empresas_a_atualizar[emp] = certidoes_a_atualizar

    def pega_cnpj(self):
        self.atualizar()
        for emp in self.empresas_a_atualizar:
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
                            self.empresas_a_atualizar[emp].append(cnpj_tratado)

    def pdf_para_jpg_renomear(self):
        print('CRIANDO IMAGENS:\n')
        for emp in self.empresas:
            os.chdir(f'{self.pdf_dir}/{str(emp)}')
            for pdf_file in os.listdir(f'{self.pdf_dir}/{str(emp)}'):
                if '00.MERGE' in pdf_file:
                    if not os.path.isdir(f'{self.pdf_dir}/{str(emp)}/Merge'):
                        os.makedirs(f'{self.pdf_dir}/{str(emp)}/Merge')
                        shutil.move(pdf_file, f'{self.pdf_dir}/{str(emp)}/Merge/{pdf_file}')
                    else:
                        shutil.move(pdf_file, f'{self.pdf_dir}/{str(emp)}/Merge/{pdf_file}')
            
                elif pdf_file.endswith(".pdf"):
                    pages = convert_from_path(pdf_file, 300, last_page=1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")
                    self.percentual += (25 / len(self.empresas))
                    print(f'Total de imagens criadas: {self.percentual}%')
        self.mensagem_log('\nIMAGENS CRIADAS COM SUCESSO!')
        self.percentual = 0

    def merge(self):
        pasta_de_trabalho = f'//hrg-74977/GEOF/HRG/PDPAS 2020/PAGAMENTO/{self.ano}-{self.mes}-{self.dia}'
        if os.path.exists(f'{pasta_de_trabalho}/Merge'):
            print('Já existe pasta para mesclagem na data informada')
        else:
            os.makedirs(f'{pasta_de_trabalho}/Merge')
        os.chdir(pasta_de_trabalho)
        for arquivo_pdf in os.listdir(pasta_de_trabalho):
            os.chdir(pasta_de_trabalho)
            if arquivo_pdf.endswith(".pdf"):
                for emp in self.empresas:
                    validação_de_partes_do_nome =[]
                    retira_espaço_empresa = emp.replace(' ', '-')
                    nome_separado = retira_espaço_empresa.split('-')
                    retira_espaço_do_arquivo = arquivo_pdf.replace(' ','-')
                    arquivo_separado = retira_espaço_do_arquivo.split('-')
                    for parte_do_nome in nome_separado:
                        contador = 0
                        if nome_separado[contador] == arquivo_separado[contador + 1]:
                            validação_de_partes_do_nome.append('OK')
                            contador += 1
                        else:
                            validação_de_partes_do_nome.append('falha')
                            contador += 1
                    print(nome_separado)
                    print(arquivo_separado)
                    print(validação_de_partes_do_nome)
                    if 'falha' not in validação_de_partes_do_nome:
                        print(emp)
                        print(arquivo_pdf.split())
                        pdf_temporário = PyPDF2.PdfFileWriter()
                        print(arquivo_pdf)
                        pagamento = open(arquivo_pdf, 'rb')
                        pagamento_lido = PyPDF2.PdfFileReader(pagamento, strict=False)
                        for página in range(pagamento_lido.numPages):
                            objeto_pagina = pagamento_lido.getPage(página)
                            pdf_temporário.addPage(objeto_pagina)
                        pasta_da_empresa = f'//hrg-74977/GEOF/CERTIDÕES/Pagamentos/{self.ano}-{self.mes}-{self.dia}/{emp}'
                        os.chdir(pasta_da_empresa)
                        for arquivo_certidão in os.listdir(pasta_da_empresa):
                            if '00.MERGE' not in arquivo_certidão:
                                certidão = open(arquivo_certidão, 'rb')
                                certidão_lida = PyPDF2.PdfFileReader(certidão)
                                for página_da_certidão in range(certidão_lida.numPages):
                                    objeto_pagina_da_certidão = certidão_lida.getPage(página_da_certidão)
                                    pdf_temporário.addPage(objeto_pagina_da_certidão)
                        compilado = open(f'{pasta_de_trabalho}/Merge/{arquivo_pdf[0:-4]}_mesclado.pdf','wb')
                        pdf_temporário.write(compilado)
                        compilado.close()
                        pagamento.close()
                        certidão.close()



    def gera_nome(self):
        print('\nRENOMEANDO CERTIDÕES:\n\n')
        for emp in self.empresas:
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
                             'GOVERNO DO DISTRITO FEDERAL': 'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                             '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                                             '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)'}
                    datas2 = {'GOVERNO DO DISTRITO FEDERAL': 'Válida até (\d) de (janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                                             '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)?(Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                             '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)'}
                    for frase in padroes:
                        if frase in certidao:
                            self.percentual += (25 / len(self.empresas))
                            print(f'{emp} - certidão {valores[frase]} renomeada - Total executado: {self.percentual}%\n')
                            if frase == 'GOVERNO DO DISTRITO FEDERAL':
                                try:
                                    data = re.compile(datas2[frase])
                                    procura = data.search(certidao)
                                    datanome = procura.group()
                                    separa = datanome.split('/')
                                    junta = '-'.join(separa)
                                except AttributeError:
                                    data = re.compile(datas[frase])
                                    procura = data.search(certidao)
                                    datanome = procura.group()
                                    separa = datanome.split('/')
                                    junta = '-'.join(separa)
                            else:
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

    def apaga_imagem(self):
        for emp in self.empresas:
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
        self.listar_cnpjs()
        padrão_cnpj = re.compile('(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)')
        validação_de_cnpj = padrão_cnpj.search(certidao).group()
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
        self.mensagem_log_sem_data(f'   UNIÃO - emissão: {emissao}; válida até: {vencimento}')
        if validação_de_cnpj in self.lista_de_cnpj:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à empresa {self.lista_de_cnpj[validação_de_cnpj]}\n')
        else:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, não possui correspondência\n')
        return (data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento), validação_de_cnpj

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
        self.listar_cnpjs()
        padrão_cnpj = re.compile('(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)')
        validação_de_cnpj = padrão_cnpj.search(certidao).group()
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
        self.mensagem_log_sem_data(f'   TST   - emissão: {emissao}; válida até: {vencimento}')
        if validação_de_cnpj in self.lista_de_cnpj:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à empresa {self.lista_de_cnpj[validação_de_cnpj]}\n')
        else:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, não possui correspondência\n')
        return (data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento), validação_de_cnpj

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
        self.listar_cnpjs()
        padrão_cnpj = re.compile('(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)')
        validação_de_cnpj = padrão_cnpj.search(certidao).group()
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
        self.mensagem_log_sem_data(f'   FGTS  - emissão: {emissao}; válida até: {vencimento}')
        if validação_de_cnpj in self.lista_de_cnpj:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à empresa {self.lista_de_cnpj[validação_de_cnpj]}\n')
        else:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, não possui correspondência\n')
        return (data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento), validação_de_cnpj

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
        self.listar_cnpjs()
        padrão_cnpj = re.compile('(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)')
        validação_de_cnpj = padrão_cnpj.search(certidao).group()
        texto = []
        meses = {'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04', 'Maio': '05', 'Junho': '06',
                 'Julho': '07', 'Agosto': '08', 'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'}
        meses2 = {'janeiro': '01', 'fevereiro': '02', 'março': '03', 'abril': '04', 'maio': '05', 'junho': '06',
                 'julho': '07', 'agosto': '08', 'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12', 'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04', 'Maio': '05', 'Junho': '06',
                 'Julho': '07', 'Agosto': '08', 'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'}
        if "GOVERNO" not in certidao:
            padrao = re.compile('Brasília, (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                                '(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                                             '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)')
            emissao_string = padrao.search(certidao)
            datasplit = [emissao_string.group().split()[1], meses[emissao_string.group().split()[3]],
                         emissao_string.group().split()[5]]
            texto.append('/'.join(datasplit))
            padrao = re.compile(
                'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                '(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                                             '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)')
            vencimento_string = padrao.search(certidao)
            datasplit2 = [vencimento_string.group().split()[2], meses[vencimento_string.group().split()[4]],
                          vencimento_string.group().split()[6]]
            texto.append('/'.join(datasplit2))
        else:
            padrao = re.compile('Certidão emitida via internet em (\d\d)/(\d\d)/(\d\d\d\d)')
            emissao_string = padrao.search(certidao)
            texto.append(emissao_string.group().split()[5])
            try:
                padrao = re.compile('Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                '(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?(julho)?(agosto)?'
                                    '(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)')
                vencimento_string = padrao.search(certidao)
                datasplit2 = [vencimento_string.group().split()[2], meses2[vencimento_string.group().split()[4]],
                             vencimento_string.group().split()[6]]
            except AttributeError:
                padrao = re.compile('Válida até (\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                '(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?(julho)?(agosto)?'
                                    '(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)')
                vencimento_string = padrao.search(certidao)
                datasplit2 = [vencimento_string.group().split()[2], meses2[vencimento_string.group().split()[4]],
                             vencimento_string.group().split()[6]]
            texto.append('/'.join(datasplit2))
        emissao = texto[0]
        vencimento = 0
        if len(texto[1]) != 10:
            vencimento = f'0{texto[1]}'
        else:
            vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        self.mensagem_log_sem_data(f'   GDF   - emissão: {emissao}; válida até: {vencimento}')
        if validação_de_cnpj in self.lista_de_cnpj:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à empresa {self.lista_de_cnpj[validação_de_cnpj]}\n')
        else:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, não possui correspondência\n')
        return (data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento), validação_de_cnpj