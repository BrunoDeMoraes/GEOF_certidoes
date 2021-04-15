from tkinter import *
from tkinter import filedialog
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
from tkinter import messagebox
import sqlite3


class Certidao:
    def __init__(self, dia, mes, ano):
        self.dia = dia
        self.mes = mes
        self.ano = ano
        self.lista_de_urls = []
        self.urls()
        self.caminho_xls = self.lista_de_urls[0][1]
        self.wb = openpyxl.load_workbook(self.caminho_xls)
        self.checagem_de_planilhas()
        self.pag = self.wb['PAGAMENTO']
        self.forn = self.wb['FORNECEDORES']
        self.listareferencia = []
        self.referencia = 0
        self.datapag = f'CERTIDÕES PARA {self.dia}/{self.mes}/{self.ano}'
        self.empresas = {}
        self.pasta_de_certidões = self.lista_de_urls[1][1]
        self.percentual = 0
        self.lista_de_cnpj = {}
        self.lista_de_cnpj_exceções = {}
        self.orgaos = ['UNIÃO', 'TST', 'FGTS', 'GDF']
        self.empresasdic = {}
        self.empresas_a_atualizar = {}
        self.caminho_de_log = f'{self.lista_de_urls[2][1]}/{self.ano}-{self.mes}-{self.dia}.txt'
        self.comprovantes_de_pagamento = f'{self.lista_de_urls[3][1]}/{self.ano}-{self.mes}-{self.dia}'
        self.certidões_para_pagamento = f'{self.lista_de_urls[4][1]}/{self.ano}-{self.mes}-{self.dia}'

    def __file__(self):
        caminho_py = __file__
        caminho_do_dir = caminho_py.split('\\')
        caminho_uso = ('/').join(caminho_do_dir[0:-1])
        return caminho_uso

    def checagem_de_planilhas(self):
        try:
            self.wb['PAGAMENTO'] and self.wb['FORNECEDORES']
        except KeyError:
            messagebox.showerror('Esse arquivo não rola!', 'O arquivo xlsx selecionado como fonte não possui as'
                                                           ' planilhas necessárias para o processamento solicitado.'
                                                           '\n\nClique em Configurações>>Caminhos>>Fonte de dados XLSX e '
                                                           'selecione um arquivo xlsx que atenda aos critérios necessários '
                                                           'para o processamento.')

    def mensagem_log(self, mensagem):
        with open(self.caminho_de_log,
                  'a') as log:
            momento = datetime.datetime.now()
            log.write(f"{mensagem} - {momento.strftime('%d/%m/%Y %H:%M:%S')}\n")
            print(f"{mensagem} - {momento.strftime('%d/%m/%Y %H:%M:%S')}")

    def mensagem_log_sem_data(self, mensagem):
        with open(self.caminho_de_log,
                  'a') as log:
            momento = datetime.datetime.now()
            log.write(f"{mensagem} - {momento.strftime('%H:%M:%S')}\n")
            print(f"{mensagem} - {momento.strftime('%H:%M:%S')}")

    def mensagem_log_sem_horario(self, mensagem):
        with open(self.caminho_de_log,
                  'a') as log:
            log.write(f"{mensagem}\n")
            print(f"{mensagem}")

    def urls(self):
        conexao = sqlite3.connect(f'{self.__file__()}/caminhos.db')
        direcionador = conexao.cursor()
        direcionador.execute("SELECT *, oid FROM urls")
        self.lista_de_urls = direcionador.fetchall()
        conexao.close()

    def pega_referencia(self):
        if os.path.exists(f'{self.comprovantes_de_pagamento}'):
            print('\nPasta para inclusão de arquivos de pagamento localizada.\n')
        else:
            os.makedirs(f'{self.comprovantes_de_pagamento}')
            print('\nPasta para inclusão de arquivos de pagamento criada com sucesso.\n')
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
            self.mensagem_log('\nA data informada não foi encontrada na lista de datas para pagamento ou não existe!')
            messagebox.showerror('Me ajuda a te ajudar!',
                                 'A data informada não foi encontrada na lista de datas para pagamento ou não existe!')
            raise Exception('A data informada não foi encontrada na lista de datas para pagamento ou não existe!!')
        elif len(self.listareferencia) > 1:
            self.mensagem_log(f'''Data informada em multiplicidade
            A data especificada foi encontrada nas células {self.listareferencia} da planilha de pagamentos: \\\hrg-74977\\GEOF\\CERTIDÕES\\Análise\\atual.xlsx.
                            \nApague as células informadas com valores duplicados e execute o programa novamente.''')
            messagebox.showerror('Me ajuda a te ajudar!',
                                 f'A data especificada foi encontrada nas células {self.listareferencia} da planilha de pagamentos: \\\hrg-74977\GEOF\CERTIDÕES\Análise\\atual.xlsx.'
                                 f'\nApague as células informadas com valores duplicados e execute o programa novamente.')
            raise Exception(
                f'A data especificada foi encontrada nas células {self.listareferencia} da planilha de pagamentos: \\\hrg-74977\GEOF\CERTIDÕES\Análise\\atual.xlsx.'
                f'\nApague as células informadas com valores duplicados e execute o programa novamente.')
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
                        if self.forn[f'M{celula.row}'].value == None:
                            continue
                        else:
                            self.empresas[emp].append(self.forn[f'M{celula.row}'].value)
                            cnpj_matriz_formatado = self.empresas[emp][3]
                            cnpj_matriz_tratado = ''
                            for digito in cnpj_matriz_formatado:
                                if digito in '0123456789':
                                    cnpj_matriz_tratado += digito
                            self.empresas[emp].append(cnpj_matriz_tratado)
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

    def listar_cnpjs_exceções(self):
        for emp in self.empresas:
            for linha in self.forn['M6':'M500']:
                for celula in linha:
                    if celula.value == None:
                        continue
                    else:
                        nome_da_empresa = self.forn[f'A{celula.row}'].value.split()
                        self.lista_de_cnpj_exceções[celula.value] = ' '.join(
                            nome_da_empresa[0:len(nome_da_empresa) - 1])
        return self.lista_de_cnpj_exceções

    def cria_diretorio(self):
        novos_dir = []
        for emp in self.empresas:
            if os.path.isdir(f'{self.pasta_de_certidões}/{str(emp)}'):
                continue
            else:
                os.makedirs(f'{self.pasta_de_certidões}/{str(emp)}/Vencidas')
                os.makedirs(f'{self.pasta_de_certidões}/{str(emp)}/Imagens')
                novos_dir.append(emp)
        self.mensagem_log(f'\nNúmero de novas pastas criadas: {len(novos_dir)} - {novos_dir}.')

    def cria_certidoes_para_pagamento(self):
        if os.path.exists(f'{self.certidões_para_pagamento}'):
            print('Já existe pasta contendo certidões para pagamento na data informada.')
            messagebox.showwarning('FICA CALMO!!!',
                                   f'''Já existe pasta contendo certidões para pagamento na data informada!

Se deseja fazer nova transferência apague o diretório:
{self.certidões_para_pagamento}''')
        else:
            os.makedirs(self.certidões_para_pagamento)
            for emp in self.empresas:
                pasta_da_empresa = f'{self.pasta_de_certidões}/{str(emp)}'
                os.makedirs(f'{self.certidões_para_pagamento}/{emp}')
                os.chdir(f'{pasta_da_empresa}')
                for pdf_file in os.listdir(f'{pasta_da_empresa}'):
                    if pdf_file.endswith(".pdf"):
                        shutil.copy(f'{pasta_da_empresa}/{pdf_file}',
                                    f'{self.certidões_para_pagamento}/{emp}/{pdf_file}')
            self.mensagem_log_sem_horario(
                f'As certidões referentes ao pagamento com data limite para a data de {self.dia}/{self.mes}/{self.ano} foram transferidas para respectiva pasta de pagamento.')
            messagebox.showinfo('Transferiu, miserávi!',
                                'As certidões que validam o pagamento foram transferidas com sucesso!')

    def certidoes_n_encontradas(self):
        total_faltando = 0
        for emp in self.empresas:
            itens = []
            faltando = []
            os.chdir(f'{self.pasta_de_certidões}/{str(emp)}')
            for item in os.listdir(f'{self.pasta_de_certidões}/{str(emp)}'):
                itens.append(item.split()[0])
            for orgao in self.orgaos:
                if orgao not in itens:
                    faltando.append(orgao)
            if faltando != []:
                try:
                    self.empresas[emp][2]
                except:
                    messagebox.showerror('Problema com o xlsx', 'O arquivo fonte de dados XLSX parece não ter sido atualizado corretamente.\n\n'
                                                                                'Tente atualizar a planilha FORNECEDORES usando a oção de colagem  que insere apenas "Valores"')
                self.mensagem_log(f'Para a empresa {emp} não foram encontradas as certidões {faltando} - CNPJ: {self.empresas[emp][2]}')
                total_faltando += 1
        if total_faltando != 0:
            self.mensagem_log(f'Adicione as certidões às respectivas pastas informadas e execute novamente o programa.')
            messagebox.showerror('Tá faltando coisa, mano!', f'''Algumas certidões não foram encontradas!
Consulte o arquivo de log, resolva as pendências indicadas e então execute novamente a análise.''')
            raise Exception(f'Adicione as certidões às respectivas pastas informadas e execute novamente o programa.')

    def pdf_para_jpg(self):
        for emp in self.empresas:
            os.chdir(f'{self.pasta_de_certidões}/{str(emp)}')
            for pdf_file in os.listdir(f'{self.pasta_de_certidões}/{str(emp)}'):
                if '00.MERGE' in pdf_file:
                    if not os.path.isdir(f'{self.pasta_de_certidões}/{str(emp)}/Merge'):
                        os.makedirs(f'{self.pasta_de_certidões}/{str(emp)}/Merge')
                        shutil.move(pdf_file, f'{self.pasta_de_certidões}/{str(emp)}/Merge/{pdf_file}')
                    else:
                        shutil.move(pdf_file, f'{self.pasta_de_certidões}/{str(emp)}/Merge/{pdf_file}')
                if pdf_file.endswith(".pdf") and pdf_file.split()[0] in self.orgaos:
                    pages = convert_from_path(pdf_file, 300, last_page=1)
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
                val, cnpj_para_comparação = objeto.confere_data(certidao, emp)
                try:
                    self.empresas[emp][1]
                except IndexError:
                    messagebox.showerror('Dados do fornecedor estão zuados!', f'Não foi possível localizar o CNPJ da '
                                                                              f'empresa {emp} na planilha FORNECEDORES'
                                                                              f' do arquivo: {self.caminho_xls}.\n\n'
                                                                              f'Verifique se há registro de CNPJ para a'
                                                                              f' empresa ou se o nome informado na'
                                                                              f' planilha PAGAMENTO é idêntico ao '
                                                                              f'inserido na planilha FORNECEDORES.')
                    self.mensagem_log(f'Não foi possível localizar o CNPJ da '
                                      f'empresa {emp} na planilha FORNECEDORES'
                                      f' do arquivo: {self.caminho_xls}.\n\n'
                                      f'Verifique se há registro de CNPJ para a'
                                      f' empresa ou se o nome informado na'
                                      f' planilha PAGAMENTO é idêntico ao '
                                      f'inserido na planilha FORNECEDORES.')
                    raise Exception(
                        f'''Não foi possível localizar o CNPJ da empresa {emp} na planilha FORNECEDORES do arquivo:
{self.caminho_xls}.
Verifique se há registro de CNPJ para a empresa ou se o nome informado na planilha PAGAMENTO é idêntico ao inserido na planilha FORNECEDORES.''')

                if len(self.empresas[emp]) > 3:
                    if val == True and cnpj_para_comparação == self.empresas[emp][3]:
                        empresadic[self.orgaos[index]] = 'OK-MATRIZ'
                    elif val == True and cnpj_para_comparação == self.empresas[emp][1]:
                        empresadic[self.orgaos[index]] = 'OK'
                    elif cnpj_para_comparação != self.empresas[emp][1] and cnpj_para_comparação != self.empresas[emp][
                        3]:
                        empresadic[self.orgaos[index]] = 'CNPJ-ERRO'
                    else:
                        empresadic[self.orgaos[index]] = 'INCOMPATÍVEL'
                else:
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
        print(
            '\n===================================================================================================\n\n'
            'Criando imagens:\n')
        for emp in self.empresas:
            os.chdir(f'{self.pasta_de_certidões}/{str(emp)}')
            for pdf_file in os.listdir(f'{self.pasta_de_certidões}/{str(emp)}'):
                if '00.MERGE' in pdf_file:
                    if not os.path.isdir(f'{self.pasta_de_certidões}/{str(emp)}/Merge'):
                        os.makedirs(f'{self.pasta_de_certidões}/{str(emp)}/Merge')
                        shutil.move(pdf_file, f'{self.pasta_de_certidões}/{str(emp)}/Merge/{pdf_file}')
                    else:
                        shutil.move(pdf_file, f'{self.pasta_de_certidões}/{str(emp)}/Merge/{pdf_file}')

                elif pdf_file.endswith(".pdf"):
                    pages = convert_from_path(pdf_file, 300, last_page=1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")
                    self.percentual += (25 / len(self.empresas))
                    print(f'Total de imagens criadas: {self.percentual}%')
        self.mensagem_log('\nImagens criadas com sucesso!')
        self.percentual = 0

    def gera_nome(self):
        print('\nRenomeando certidões:\n\n')
        for emp in self.empresas:
            os.chdir(f'{self.pasta_de_certidões}/{(emp)}')
            origem = f'{self.pasta_de_certidões}/{emp}'
            for imagem in os.listdir(origem):
                if imagem.endswith(".jpg"):
                    certidao = pytesseract.image_to_string(Image.open(f'{origem}/{imagem}'), lang='por')
                    padroes = ['FGTS - CRF', 'Brasília,', 'JUSTIÇA DO TRABALHO', 'MINISTÉRIO DA FAZENDA',
                               'GOVERNO DO DISTRITO FEDERAL']
                    valores = {'FGTS - CRF': 'FGTS', 'Brasília,': 'GDF', 'JUSTIÇA DO TRABALHO': 'TST',
                               'MINISTÉRIO DA FAZENDA': 'UNIÃO', 'GOVERNO DO DISTRITO FEDERAL': 'GDF'}
                    datas = {'FGTS - CRF': 'a (\d\d)/(\d\d)/(\d\d\d\d)',
                             'Brasília,': 'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                          '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)',
                             'JUSTIÇA DO TRABALHO': 'Validade: (\d\d)/(\d\d)/(\d\d\d\d)',
                             'MINISTÉRIO DA FAZENDA': 'Válida até (\d\d)/(\d\d)/(\d\d\d\d)',
                             'GOVERNO DO DISTRITO FEDERAL': 'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                                            '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                                                            '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)'}
                    datas2 = {
                        'GOVERNO DO DISTRITO FEDERAL': 'Válida até (\d) de (janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                                                       '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)?(Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                                       '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)'}
                    for frase in padroes:
                        if frase in certidao:
                            self.percentual += (25 / len(self.empresas))
                            print(
                                f'{emp} - certidão {valores[frase]} renomeada - Total executado: {self.percentual}%\n')
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
        print('\nProcesso de renomeação de certidões executado com sucesso!')

    def merge(self):
        if os.path.exists(f'{self.comprovantes_de_pagamento}/Mesclados'):
            print('Já existe pasta para mesclagem na data informada')
            messagebox.showwarning('FICA CALMO!!!', f'''Já existe pasta para mesclagem na data informada!

Se deseja fazer nova mesclagem apague o diretório:
{self.comprovantes_de_pagamento}/Mesclados.''')
        else:
            os.makedirs(f'{self.comprovantes_de_pagamento}/Mesclados')
            os.chdir(self.comprovantes_de_pagamento)
            for arquivo_pdf in os.listdir(self.comprovantes_de_pagamento):
                os.chdir(self.comprovantes_de_pagamento)
                if arquivo_pdf.endswith(".pdf"):
                    for emp in self.empresas:
                        validação_de_partes_do_nome = []
                        retira_espaço_empresa = emp.replace(' ', '-')
                        nome_separado = retira_espaço_empresa.split('-')
                        retira_espaço_do_arquivo = arquivo_pdf.replace(' ', '-')
                        arquivo_separado = retira_espaço_do_arquivo.split('-')
                        for parte_do_nome in nome_separado:
                            contador = 0
                            if nome_separado[contador] == arquivo_separado[contador + 1]:
                                validação_de_partes_do_nome.append('OK')
                                contador += 1
                            else:
                                validação_de_partes_do_nome.append('falha')
                                contador += 1
                        if 'falha' not in validação_de_partes_do_nome:
                            print(emp)
                            pdf_temporário = PyPDF2.PdfFileWriter()
                            pagamento = open(arquivo_pdf, 'rb')
                            try:
                                pagamento_lido = PyPDF2.PdfFileReader(pagamento, strict=False)
                            except:
                                messagebox.showerror('Arquivo zuado!!!', f"o arquivo {arquivo_pdf} está corrompido")
                            for página in range(pagamento_lido.numPages):
                                objeto_pagina = pagamento_lido.getPage(página)
                                pdf_temporário.addPage(objeto_pagina)
                            pasta_da_empresa = f'{self.certidões_para_pagamento}/{emp}'
                            os.chdir(pasta_da_empresa)
                            for arquivo_certidão in os.listdir(pasta_da_empresa):
                                if '00.MERGE' not in arquivo_certidão:
                                    certidão = open(arquivo_certidão, 'rb')
                                    certidão_lida = PyPDF2.PdfFileReader(certidão)
                                    for página_da_certidão in range(certidão_lida.numPages):
                                        objeto_pagina_da_certidão = certidão_lida.getPage(página_da_certidão)
                                        pdf_temporário.addPage(objeto_pagina_da_certidão)
                            compilado = open(
                                f'{self.comprovantes_de_pagamento}/Mesclados/{arquivo_pdf[0:-4]}_mesclado.pdf', 'wb')
                            pdf_temporário.write(compilado)
                            compilado.close()
                            pagamento.close()
                            certidão.close()
            messagebox.showinfo('Mesclou, miserávi!!!',
                                'Digitalizações de pagamentos e respectivas certidões mescladas com sucesso!')

    def apaga_imagem(self):
        for emp in self.empresas:
            if not os.path.exists(f'{self.pasta_de_certidões}/{str(emp)}'):
                messagebox.showerror('Tem essa pasta aí não, locão!',
                                     f'A pasta {self.pasta_de_certidões}/{str(emp)} ainda não existe.\n\n'
                                     f'Antes de tentar renomear as certidões, execute a opção "Analisar certidões". '
                                     f'A referida opção criará as pastas necessárias e indicará o que '
                                     f'precisa ser atualizado antes do processo de renomeação.')
            os.chdir(f'{self.pasta_de_certidões}/{str(emp)}')
            for arquivo in os.listdir(f'{self.pasta_de_certidões}/{str(emp)}'):
                if arquivo.endswith(".jpg"):
                    os.unlink(f'{self.pasta_de_certidões}/{str(emp)}/{arquivo}')


class Uniao(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(f'{self.pasta_de_certidões}/{str(emp)}')
        for imagem in os.listdir(f'{self.pasta_de_certidões}/{str(emp)}'):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'UNIÃO':
                certidao = pytesseract.image_to_string(
                    Image.open(f'{self.pasta_de_certidões}/{emp}/{imagem}'),
                    lang='por')
                return certidao

    def confere_data(self, certidao, emp):
        self.listar_cnpjs()
        self.listar_cnpjs_exceções()
        padrão_cnpj = re.compile('(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)')
        try:
            validação_de_cnpj = padrão_cnpj.search(certidao).group()
        except AttributeError:
            self.mensagem_log(
                f'Execução interrompida!!!\nNão foi possível encontrar o padrão de CNPJ na certidão UNIÃO da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
            messagebox.showerror('Esse arquivo não rola!',
                                 f'''Não foi possível encontrar o padrão de CNPJ na certidão UNIÃO da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão UNIÃO da empresa {emp} inválido.')
        texto = []
        padrao = re.compile('do dia (\d\d)/(\d\d)/(\d\d\d\d)')
        emissao_string = padrao.search(certidao)
        try:
            texto.append(emissao_string.group().split()[2])
        except AttributeError:
            self.mensagem_log(
                f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de emissão na certidão UNIÃO da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
            messagebox.showerror('Esse arquivo não rola!',
                                 f'''Não foi possível encontrar o padrão de data de emissão na certidão UNIÃO da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão UNIÃO da empresa {emp} inválido.')

        padrao = re.compile('Válida até (\d\d)/(\d\d)/(\d\d\d\d)')
        vencimento_string = padrao.search(certidao)
        try:
            texto.append(vencimento_string.group().split()[2])
        except AttributeError:
            self.mensagem_log(
                f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de vencimento na certidão UNIÃO da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
            messagebox.showerror('Esse arquivo não rola!',
                                 f'''Não foi possível encontrar o padrão de data de vencimento na certidão UNIÃO da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão UNIÃO da empresa {emp} inválido.')
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        self.mensagem_log_sem_data(f'   UNIÃO - emissão: {emissao}; válida até: {vencimento}')
        if validação_de_cnpj in self.lista_de_cnpj:
            self.mensagem_log_sem_horario(
                f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à empresa {self.lista_de_cnpj[validação_de_cnpj]}\n')
        elif validação_de_cnpj in self.lista_de_cnpj_exceções:
            self.mensagem_log_sem_horario(
                f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à matriz da empresa {self.lista_de_cnpj_exceções[validação_de_cnpj]}\n')
        else:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, não possui correspondência\n')
        return (data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento), validação_de_cnpj


class Tst(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(f'{self.pasta_de_certidões}/{str(emp)}')
        for imagem in os.listdir(f'{self.pasta_de_certidões}/{str(emp)}'):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'TST':
                certidao = pytesseract.image_to_string(
                    Image.open(f'{self.pasta_de_certidões}/{emp}/{imagem}'),
                    lang='por')
                return certidao

    def confere_data(self, certidao, emp):
        self.listar_cnpjs()
        padrão_cnpj = re.compile('(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)')
        try:
            validação_de_cnpj = padrão_cnpj.search(certidao).group()
        except AttributeError:
            self.mensagem_log(
                f'Execução interrompida!!!\nNão foi possível encontrar o padrão de CNPJ na certidão TST da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
            messagebox.showerror('Esse arquivo não rola!',
                                 f'''Não foi possível encontrar o padrão de CNPJ na certidão TST da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão TST da empresa {emp} inválido.')
        texto = []
        padrao = re.compile('Expedição: (\d\d)/(\d\d)/(\d\d\d\d)')
        emissao_string = padrao.search(certidao)
        try:
            texto.append(emissao_string.group().split()[1])
        except AttributeError:
            self.mensagem_log(
                f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de emissão na certidão TST da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
            messagebox.showerror('Esse arquivo não rola!',
                                 f'''Não foi possível encontrar o padrão de data de emissão na certidão TST da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão TST da empresa {emp} inválido.')
        padrao = re.compile('Validade: (\d\d)/(\d\d)/(\d\d\d\d)')
        vencimento_string = padrao.search(certidao)
        try:
            texto.append(vencimento_string.group().split()[1])
        except AttributeError:
            self.mensagem_log(
                f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de vencimento na certidão TST da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
            messagebox.showerror('Esse arquivo não rola!',
                                 f'''Não foi possível encontrar o padrão de data de vencimento na certidão TST da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão TST da empresa {emp} inválido.')
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        self.mensagem_log_sem_data(f'   TST   - emissão: {emissao}; válida até: {vencimento}')
        if validação_de_cnpj in self.lista_de_cnpj:
            self.mensagem_log_sem_horario(
                f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à empresa {self.lista_de_cnpj[validação_de_cnpj]}\n')
        else:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, não possui correspondência\n')
        return (data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento), validação_de_cnpj


class Fgts(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(f'{self.pasta_de_certidões}/{str(emp)}')
        for imagem in os.listdir(f'{self.pasta_de_certidões}/{str(emp)}'):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'FGTS':
                certidao = pytesseract.image_to_string(
                    Image.open(f'{self.pasta_de_certidões}/{emp}/{imagem}'),
                    lang='por')
                return certidao

    def confere_data(self, certidao, emp):
        self.listar_cnpjs()
        padrão_cnpj = re.compile('(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)')
        try:
            validação_de_cnpj = padrão_cnpj.search(certidao).group()
        except AttributeError:
            self.mensagem_log(
                f'Execução interrompida!!!\nNão foi possível encontrar o padrão de CNPJ na certidão FGTS da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
            messagebox.showerror('Esse arquivo não rola!',
                                 f'''Não foi possível encontrar o padrão de CNPJ na certidão FGTS da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão FGTS da empresa {emp} inválido.')
        texto = []
        padrao = re.compile('(\d\d)/(\d\d)/(\d\d\d\d) a (\d\d)/(\d\d)/(\d\d\d\d)')
        emissao_string = padrao.search(certidao)
        try:
            texto.append(emissao_string.group().split()[0])
        except AttributeError:
            self.mensagem_log(
                f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de emissão e vencimento na certidão FGTS da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
            messagebox.showerror('Esse arquivo não rola!',
                                 f'''Não foi possível encontrar o padrão de data de emissão e vencimento na certidão FGTS da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão FGTS da empresa {emp} inválido.')
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
            self.mensagem_log_sem_horario(
                f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à empresa {self.lista_de_cnpj[validação_de_cnpj]}\n')
        else:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, não possui correspondência\n')
        return (data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento), validação_de_cnpj


class Gdf(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(f'{self.pasta_de_certidões}/{str(emp)}')
        for imagem in os.listdir(f'{self.pasta_de_certidões}/{str(emp)}'):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'GDF':
                certidao = pytesseract.image_to_string(
                    Image.open(f'{self.pasta_de_certidões}/{emp}/{imagem}'),
                    lang='por')
                return certidao

    def confere_data(self, certidao, emp):
        self.listar_cnpjs()
        padrão_cnpj = re.compile('(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)')
        try:
            validação_de_cnpj = padrão_cnpj.search(certidao).group()
        except AttributeError:
            self.mensagem_log(f'''Execução interrompida!!!
Não foi possível encontrar o padrão de CNPJ na certidão GDF da empresa {emp}.
O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            messagebox.showerror('Esse arquivo não rola!',
                                 f'''Não foi possível encontrar o padrão de CNPJ na certidão GDF da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão GDF da empresa {emp} inválido.')
        texto = []
        meses = {'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04', 'Maio': '05', 'Junho': '06',
                 'Julho': '07', 'Agosto': '08', 'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'}
        meses2 = {'janeiro': '01', 'fevereiro': '02', 'março': '03', 'abril': '04', 'maio': '05', 'junho': '06',
                  'julho': '07', 'agosto': '08', 'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12',
                  'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04', 'Maio': '05', 'Junho': '06',
                  'Julho': '07', 'Agosto': '08', 'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'}
        if "GOVERNO" not in certidao:
            padrao = re.compile(
                'Brasília, (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                '(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)')
            emissao_string = padrao.search(certidao)
            try:
                datasplit = [emissao_string.group().split()[1], meses[emissao_string.group().split()[3]],
                             emissao_string.group().split()[5]]
            except AttributeError:
                self.mensagem_log(
                    f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de emissão na certidão GDF da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
                messagebox.showerror('Esse arquivo não rola!',
                                     f'''Não foi possível encontrar o padrão de data de emissão na certidão GDF da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
                raise Exception(f'Arquivo da certidão GDF da empresa {emp} inválido.')

            texto.append('/'.join(datasplit))
            padrao = re.compile(
                'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                '(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)')
            vencimento_string = padrao.search(certidao)
            try:
                datasplit2 = [vencimento_string.group().split()[2], meses[vencimento_string.group().split()[4]],
                              vencimento_string.group().split()[6]]
            except AttributeError:
                self.mensagem_log(
                    f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de vencimento na certidão GDF da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
                messagebox.showerror('Esse arquivo não rola!',
                                     f'''Não foi possível encontrar o padrão de data de vencimento na certidão GDF da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
                raise Exception(f'Arquivo da certidão GDF da empresa {emp} inválido.')

            texto.append('/'.join(datasplit2))
        else:
            padrao = re.compile('Certidão emitida via internet em (\d\d)/(\d\d)/(\d\d\d\d)')
            emissao_string = padrao.search(certidao)
            try:
                texto.append(emissao_string.group().split()[5])
            except AttributeError:
                self.mensagem_log(
                    f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de emissão na certidão GDF da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
                messagebox.showerror('Esse arquivo não rola!',
                                     f'''Não foi possível encontrar o padrão de data de emissão na certidão GDF da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
                raise Exception(f'Arquivo da certidão GDF da empresa {emp} inválido.')

            padrao = re.compile(
                'Válida até (\d)?(\d\d)? de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                '(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?(julho)?(agosto)?'
                '(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)')
            vencimento_string = padrao.search(certidao)
            try:
                datasplit2 = [vencimento_string.group().split()[2], meses2[vencimento_string.group().split()[4]],
                              vencimento_string.group().split()[6]]
            except AttributeError:
                self.mensagem_log(
                    f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de vencimento na certidão GDF da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
                messagebox.showerror('Esse arquivo não rola!',
                                     f'''Não foi possível encontrar o padrão de data de vencimento na certidão GDF da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
                raise Exception(f'Arquivo da certidão GDF da empresa {emp} inválido.')
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
            self.mensagem_log_sem_horario(
                f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à empresa {self.lista_de_cnpj[validação_de_cnpj]}\n')
        else:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, não possui correspondência\n')
        return (data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento), validação_de_cnpj


class Analisador:
    opções = [
        'Selecione uma opção', 'Renomear arquivos',
        'Renomear todos os arquivos de uma pasta', 'Renomear todas as certidões da lista de pagamento']

    def __init__(self, tela):
        self.urls = []
        self.cria_pastas_de_trabalho()
        self.cria_bd()
        self.frame_mestre = LabelFrame(tela, padx=0, pady=0)
        self.frame_mestre.pack(padx=1, pady=1)

        self.frame_data = LabelFrame(self.frame_mestre, padx=0, pady=0)

        self.frame_renomear = LabelFrame(self.frame_mestre, padx=0, pady=0)

        self.menu_certidões = Menu(tela)
        self.menu_configurações = Menu(self.menu_certidões)
        self.menu_certidões.add_cascade(label='Configurações', menu=self.menu_configurações)
        self.menu_configurações.add_separator()
        self.menu_configurações.add_command(label='Caminhos', command=self.abrir_janela_caminhos)
        self.menu_configurações.add_separator()


        self.titulo = Label(self.frame_data, text='    Indique a data limite pretendida para o próximo pagamento e em seguida escolha uma das seguintes opções:    ',
                            pady=0, padx=0, bg='green', fg='white', bd=2, relief=SUNKEN,font=('Helvetica', 10, 'bold'))

        self.dia_etiqueta = Label(self.frame_data, text='Dia', padx=22, pady=0, bg='green', fg='white', bd=2, relief=SUNKEN,
                             font=('Helvetica', 10, 'bold'))
        self.mes_etiqueta = Label(self.frame_data, text='Mês', padx=22, pady=0, bg='green', fg='white', bd=2, relief=SUNKEN,
                             font=('Helvetica', 10, 'bold'))
        self.ano_etiqueta = Label(self.frame_data, text='Ano', padx=22, pady=0, bg='green', fg='white', bd=2, relief=SUNKEN,
                             font=('Helvetica', 10, 'bold'))

        self.variavel = StringVar()
        self.variavel.set(" ")
        self.variavel2 = StringVar()
        self.variavel2.set(" ")
        self.variavel3 = StringVar()
        self.variavel3.set(" ")
        self.dias = [' ']
        self.meses = [' ']
        self.anos = [' ']

        self.cria_calendario()

        self.botao_abrir_log = Button(self.frame_data, text='Abrir log', command=self.abrir_log, padx=0,
                                     pady=0, bg='white',
                                     fg='green', font=('Helvetica', 9, 'bold'), bd=1)
        self.validacao1 = OptionMenu(self.frame_data, self.variavel, *self.dias)
        self.validacao2 = OptionMenu(self.frame_data, self.variavel2, *self.meses)
        self.validacao3 = OptionMenu(self.frame_data, self.variavel3, *self.anos)

        self.titulo_analisar = Label(self.frame_mestre, text='Utilize esta opção para identificar quais certidões devem ser atualizadas ou se há requisitos a cumprir para a devida execução da análise.', pady=0, padx=0, bg='white',
                                fg='black', font=('Helvetica', 9, 'bold'))

        self.botao_analisar = Button(self.frame_mestre, text='Analisar\ncertidões', command=self.executa, padx=30, pady=1, bg='green',
                                fg='white', font=('Helvetica', 9, 'bold'), bd=1)


        self.titulo_renomear = Label(self.frame_mestre, text='Após atualizar as certidões, selecione uma das opções para padronizar os nomes dos\narquivos e em seguida faça nova análise para certificar que está tudo OK.', pady=0, padx=0,
                                     bg='white', fg='black', font=('Helvetica', 9, 'bold'))

        self.variavel_de_opções = StringVar()
        self.variavel_de_opções.set("Selecione uma opção")
        self.validacao = OptionMenu(self.frame_mestre, self.variavel_de_opções, *Analisador.opções)

        self.arquivo_selecionado = 'Selecione os arquivos que deseja renomear'
        self.pasta_selecionada = 'Selecione a pasta que deseja renomear'


        self.botao_renomear_tudo = Button(self.frame_mestre, text='Renomear\ncertidões',
                                          command=self.selecionador_de_opções, padx=30, pady=1, bg='green', fg='white',
                                          font=('Helvetica', 9, 'bold'), bd=1)

        self.titulo_transfere_arquivos = Label(self.frame_mestre, text='Esta opção transfere as certidões que validam o pagamento para uma pasta identificada pela data.\nEsse passo deve ser executado logo após a análise definitiva antes do pagamento.', pady=0, padx=0, bg='white',
                                               fg='black',
                                               font=('Helvetica', 9, 'bold'))

        self.botao_transfere_arquivos = Button(self.frame_mestre, text='Transferir\ncertidões', command=self.transfere_certidoes,
                                          padx=30, pady=1, bg='green',
                                          fg='white', font=('Helvetica', 9, 'bold'), bd=1)

        self.titulo_mescla_arquivos = Label(self.frame_mestre, text='Após o pagamento utilize esta opção para mesclar os comprovantes de pagamento digitalizados com suas respectivas certidões.', pady=0, padx=0, bg='white', fg='black',
                                       font=('Helvetica', 9, 'bold'))

        self.botao_mescla_arquivos = Button(self.frame_mestre, text=' Mesclar  \narquivos', command=self.mescla_certidoes, padx=30,
                                       pady=1, bg='green',
                                       fg='white', font=('Helvetica', 9, 'bold'), bd=1)


        self.roda_pe = Label(self.frame_mestre, text="SRSSU/DA/GEOF   ", pady=0, padx=0, bg='green', fg='white',
                             font=('Helvetica', 8, 'italic'), anchor=E)

        self.frame_data.grid(row=0, column=1, columnspan=7, rowspan=1, pady=0, sticky=W+E)
        self.titulo.grid(row=0, column=1, columnspan=5, rowspan=1, pady=0, sticky=W+E)
        self.dia_etiqueta.grid(row=0, column=6, pady=0, ipadx=0, ipady=0)
        self.mes_etiqueta.grid(row=0, column=7, pady=0, ipadx=0, ipady=0)
        self.ano_etiqueta.grid(row=0, column=8, pady=0, ipadx=0, ipady=0)
        self.botao_abrir_log.grid(row=1, column=1, pady=0)
        self.validacao1.grid(row=1, column=6, pady=0)
        self.validacao2.grid(row=1, column=7, pady=0)
        self.validacao3.grid(row=1, column=8, pady=0)

        self.titulo_analisar.grid(row=1, column=1,  columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W+E)
        self.botao_analisar.grid(row=2, column=1, columnspan=7, padx=0, pady=10)
        self.titulo_renomear.grid(row=3, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W+E)
        self.validacao.grid(row=4, column=1, columnspan=7, padx=0, pady=10)
        self.botao_renomear_tudo.grid(row=5, column=1, columnspan=7, padx=0, pady=10)

        self.titulo_transfere_arquivos.grid(row=7, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W+E)
        self.botao_transfere_arquivos.grid(row=8, column=1, columnspan=7, padx=0, pady=10)
        self.titulo_mescla_arquivos.grid(row=9, column=1, columnspan=7, padx=0, pady=0, ipadx=0, ipady=8, sticky=W+E)
        self.botao_mescla_arquivos.grid(row=10, column=1, columnspan=7, padx=0, pady=10)

        self.roda_pe.grid(row=11, column=1, columnspan=10, pady=5, sticky=W+E)

    def __file__(self):
        caminho_py = __file__
        caminho_do_dir = caminho_py.split('\\')
        caminho_uso = ('/').join(caminho_do_dir[0:-1])
        return caminho_uso

    def cria_pastas_de_trabalho(self):
        pastas_de_trabalho = ['Certidões','Logs de conferência', 'Certidões para pagamento', 'Comprovantes de pagamento']
        for pasta in pastas_de_trabalho:
            if not os.path.exists(f'{self.__file__()}/{pasta}'):
                os.makedirs(f'{self.__file__()}/{pasta}')
                print(f'Pasta de trabalho {pasta} criada com sucesso!\n')
            else:
                print(f'Pasta de trabalho {pasta} localizada.\n')


    def abrir_janela_caminhos(self):
        conexao = sqlite3.connect(f'{self.__file__()}/caminhos.db')
        direcionador = conexao.cursor()
        direcionador.execute("SELECT *, oid FROM urls")
        self.urls = direcionador.fetchall()
        conexao.close()

        self.janela_de_caminhos = Toplevel()
        self.janela_de_caminhos.title('Lista de caminhos')
        self.janela_de_caminhos.resizable(False, False)
        self.frame_de_caminhos = LabelFrame(self.janela_de_caminhos, padx=0, pady=0)
        self.frame_de_caminhos.pack(padx=1, pady=1)
        self.botao_xlsx = Button(self.frame_de_caminhos, text='Fonte de\ndados XLSX', command=self.altera_caminho_xlsl,
                                 padx=0, pady=0, bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)
        self.caminho_xlsx = Entry(self.frame_de_caminhos, width=70)
        self.botao_pasta_de_certidões = Button(self.frame_de_caminhos, text='Pasta de\ncertidões', command=self.altera_caminho_pasta_de_certidões,
                                               padx=0, pady=0, bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)
        self.caminho_pasta_de_certidões = Entry(self.frame_de_caminhos, width=70)
        self.botao_log = Button(self.frame_de_caminhos, text='Pasta de\nlogs', command=self.altera_caminho_log, padx=0, pady=0,
                                bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)
        self.caminho_log = Entry(self.frame_de_caminhos, width=70)
        self.pasta_pagamento = Button(self.frame_de_caminhos, text='Comprovantes\nde pagamentos', command=self.altera_caminho_pasta_pagamento,
                                      padx=0, pady=0, bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)
        self.caminho_pasta_pagamento = Entry(self.frame_de_caminhos, width=70)
        self.certidões_para_pagamento = Button(self.frame_de_caminhos, text='Certidões para\npagamento', command=self.altera_caminho_certidões_para_pagamento, padx=0,
                                      pady=0, bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)
        self.caminho_certidões_para_pagamento = Entry(self.frame_de_caminhos, width=70)

        self.gravar_alterações = Button(self.frame_de_caminhos, text='Gravar alterações',
                                               command=self.atualizar_xlsx, padx=10,
                                               pady=10, bg='green', fg='white', font=('Helvetica', 8, 'bold'), bd=1)

        self.botao_xlsx.grid(row=1, column=1, columnspan=1, padx = 15, pady=10, ipadx=5, ipady=13, sticky=W+E)
        self.caminho_xlsx.insert(0, self.urls[0][1])
        self.caminho_xlsx.grid(row=1, column=2, padx=20)
        self.botao_pasta_de_certidões.grid(row=2, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_pasta_de_certidões.insert(0, self.urls[1][1])
        self.caminho_pasta_de_certidões.grid(row=2, column=2, padx=20)
        self.botao_log.grid(row=3, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_log.insert(0, self.urls[2][1])
        self.caminho_log.grid(row=3, column=2, padx=20)
        self.certidões_para_pagamento.grid(row=4, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_certidões_para_pagamento.insert(0, self.urls[4][1])
        self.caminho_certidões_para_pagamento.grid(row=4, column=2, padx=20)
        self.pasta_pagamento.grid(row=5, column=1, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13, sticky=W + E)
        self.caminho_pasta_pagamento.insert(0, self.urls[3][1])
        self.caminho_pasta_pagamento.grid(row=5, column=2, padx=20)
        self.gravar_alterações.grid(row=6, column=2, columnspan=1, padx=15, pady=10, ipadx=10, ipady=13)


    def altera_caminho_xlsl(self):
        caminho = filedialog.askopenfilename(initialdir=self.__file__(),filetypes=(('Arquivos', '*.xlsx'),
                                                                                   ("Todos os arquivos", '*.*')))
        self.caminho_xlsx.destroy()
        self.caminho_xlsx = Entry(self.frame_de_caminhos, width=30)
        self.caminho_xlsx.insert(0, caminho)
        self.caminho_xlsx.grid(row=1, column=2, padx=20)

    def altera_caminho_pasta_de_certidões(self):
        caminho = filedialog.askdirectory(initialdir=self.__file__())
        self.caminho_pasta_de_certidões.destroy()
        self.caminho_pasta_de_certidões = Entry(self.frame_de_caminhos, width=30)
        self.caminho_pasta_de_certidões.insert(0, caminho)
        self.caminho_pasta_de_certidões.grid(row=2, column=2, padx=20)

    def altera_caminho_log(self):
        caminho = filedialog.askdirectory(initialdir=self.__file__())
        self.caminho_log.destroy()
        self.caminho_log = Entry(self.frame_de_caminhos, width=30)
        self.caminho_log.insert(0, caminho)
        self.caminho_log.grid(row=3, column=2, padx=20)

    def altera_caminho_pasta_pagamento(self):
        caminho = filedialog.askdirectory(initialdir=self.__file__())
        self.caminho_pasta_pagamento.destroy()
        self.caminho_pasta_pagamento = Entry(self.frame_de_caminhos, width=30)
        self.caminho_pasta_pagamento.insert(0, caminho)
        self.caminho_pasta_pagamento.grid(row=4, column=2, padx=20)

    def altera_caminho_certidões_para_pagamento(self):
        caminho = filedialog.askdirectory(initialdir=self.__file__())
        self.caminho_certidões_para_pagamento.destroy()
        self.caminho_certidões_para_pagamento = Entry(self.frame_de_caminhos, width=30)
        self.caminho_certidões_para_pagamento.insert(0, caminho)
        self.caminho_certidões_para_pagamento.grid(row=5, column=2, padx=20)

    def atualizar_xlsx(self):
        resposta = messagebox.askyesno('Vc sabe o que está fazendo?','Tem certeza que deseja alterar a configuração dos caminhos de pastas e arquivos?')
        if resposta == True:
            conexao = sqlite3.connect(f'{self.__file__()}/caminhos.db')
            direcionador = conexao.cursor()
            direcionador.execute("UPDATE urls SET url = :caminho_xlsx WHERE oid = 1",
                                 {"caminho_xlsx": self.caminho_xlsx.get()})
            direcionador.execute("UPDATE urls SET url = :pasta_de_certidões WHERE oid = 2",
                                 {"pasta_de_certidões": self.caminho_pasta_de_certidões.get()})
            direcionador.execute("UPDATE urls SET url = :caminho_de_log WHERE oid = 3",
                                 {"caminho_de_log": self.caminho_log.get()})
            direcionador.execute("UPDATE urls SET url = :comprovantes_de_pagamento WHERE oid = 4",
                {"comprovantes_de_pagamento": self.caminho_pasta_pagamento.get()})
            direcionador.execute("UPDATE urls SET url = :certidões_para_pagamento WHERE oid = 5",
                                 {"certidões_para_pagamento": self.caminho_certidões_para_pagamento.get()})
            conexao.commit()
            conexao.close()
            self.janela_de_caminhos.destroy()
            print('\nOs caminhos para pastas e arquivos utilizados pelo sistema foram atualizados.\n')
            messagebox.showinfo('Fez porque quis!',"Caminhos para pastas e arquivos utilizados pelo sistema atualizados com sucesso!")
        else:
            self.janela_de_caminhos.destroy()

    def cria_bd(self):
        pastas_de_trabalho = ['Certidões', 'Logs de conferência', 'Certidões para pagamento',
                              'Comprovantes de pagamento']
        for pasta in pastas_de_trabalho:
            if not os.path.exists(f'{self.__file__()}/{pasta}'):
                os.makedirs(f'{self.__file__()}/{pasta}')

        if not os.path.exists(f'{self.__file__()}/caminhos.db'):
            conexao = sqlite3.connect(f'{self.__file__()}/caminhos.db')
            direcionador = conexao.cursor()
            direcionador.execute('CREATE TABLE urls (variavel text, url text)')
            caminhos = {'caminho_xlsx': f'{self.__file__()}/listas.xlsx',
                        'pasta_de_certidões': f'{self.__file__()}/Certidões',
                        'caminho_de_log': f'{self.__file__()}/Logs de conferência',
                        'comprovantes_de_pagamento': f'{self.__file__()}/Comprovantes de pagamento',
                        'certidões_para_pagamento': f'{self.__file__()}/Certidões para pagamento'}
            for caminho in caminhos:
                direcionador.execute('INSERT INTO urls VALUES (:variavel, :url)',
                                     {"variavel": caminho, "url": caminhos[caminho]})
                conexao.commit()
            direcionador.execute("SELECT *, oid FROM urls")
            self.urls = direcionador.fetchall()
            for registro in self.urls:
                print(f'{registro[0]}: {registro[1]}\n')
            conexao.close()
        else:
            print('Banco de dados localizado.')
            conexao = sqlite3.connect(f'{self.__file__()}/caminhos.db')
            direcionador = conexao.cursor()
            direcionador.execute("SELECT *, oid FROM urls")
            self.urls = direcionador.fetchall()
            for registro in self.urls:
                print(f'{registro[0]}: {registro[1]}\n')
                conexao.close()


    def abrir_log(self):
        conexao = sqlite3.connect(f'{self.__file__()}/caminhos.db')
        direcionador = conexao.cursor()
        direcionador.execute("SELECT *, oid FROM urls")
        self.urls = direcionador.fetchall()
        conexao.close()
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        if not os.path.exists(f'{self.urls[2][1]}/{ano}-{mes}-{dia}.txt') or (dia, mes, ano) == (' ', ' ', ' '):
            messagebox.showerror('Me ajuda a te ajudar!',
                                 'Não existe log para a data informada.')
        else:
            caminho = f'{self.urls[2][1]}/{ano}-{mes}-{dia}.txt'
            novo_caminho = caminho.replace('/', '\\')
            os.startfile(novo_caminho)


    def cria_calendario(self):
        contador_dia = 1
        while contador_dia <= 31:
            if contador_dia < 10:
                self.dias.append(f"0{contador_dia}")
                contador_dia += 1
            else:
                self.dias.append(str(contador_dia))
                contador_dia += 1
        contador_mes = 1
        while contador_mes <= 12:
            if contador_mes < 10:
                self.meses.append(f"0{contador_mes}")
                contador_mes += 1
            else:
                self.meses.append(str(contador_mes))
                contador_mes += 1
        contador_anos = 2010
        while contador_anos <= 2040:
            self.anos.append(str(contador_anos))
            contador_anos += 1
        return self.dias, self.meses, self.anos

    def atualiza_urls(self):
        conexao = sqlite3.connect(f'{self.__file__()}/caminhos.db')
        direcionador = conexao.cursor()
        direcionador.execute("SELECT *, oid FROM urls")
        self.urls = direcionador.fetchall()
        conexao.close()

    def checa_urls(self):
        if not os.path.exists(self.urls[0][1]):
            messagebox.showerror('Sumiu!!!',
                                 'O arquivo xlsx selecionado como fonte foi apagado, removido ou não existe.'
                                 '\n\nClique em Configurações>>Caminhos>>Fonte de dados XLSX e '
                                 'selecione um arquivo xlsx que atenda aos critérios necessários '
                                 'para o processamento.')
        elif not os.path.exists(self.urls[1][1]):
            messagebox.showerror('Sumiu!!!',
                                 'A pasta apontada como fonte para cetidões foi apagada, removida ou não existe.'
                                 '\n\nClique em Configurações>>Caminhos>>Pasta de certidões e '
                                 'selecione uma pasta que contenha as certidões que devem ser analisadas.')
        elif not os.path.exists(self.urls[2][1]):
            messagebox.showerror('Sumiu!!!',
                                 'A pasta apontada como fonte e destino para logs foi apagada, removida ou não existe.'
                                 '\n\nClique em Configurações>>Caminhos>>Pasta de logs e '
                                 'selecione uma pasta onde os logs serão criados.')
        elif not os.path.exists(self.urls[4][1]):
            messagebox.showerror('Sumiu!!!',
                                 'A pasta apontada como destino de cetidões para pagamento foi apagada, removida ou não existe.'
                                 '\n\nClique em Configurações>>Caminhos>>Cetidões para pagamento e '
                                 'selecione uma pasta para direcionar as certidões do pagamento.')
        elif not os.path.exists(self.urls[3][1]):
            messagebox.showerror('Sumiu!!!',
                                 'A pasta apontada como fonte de comprovantes de pagamento foi apagada, removida ou não existe.'
                                 '\n\nClique em Configurações>>Caminhos>>Comprovantes de pagamento e '
                                 'selecione uma pasta que contenha os comprovantes de pagamento.')

    def executa(self):
        tempo_inicial = time.time()
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        self.atualiza_urls()
        if not os.path.exists(self.urls[0][1]) or not os.path.exists(self.urls[1][1]) or not os.path.exists(self.urls[2][1])\
                or not os.path.exists(self.urls[4][1]) or not os.path.exists(self.urls[3][1]):
            self.checa_urls()
        else:
            obj1 = Certidao(dia, mes, ano)

            obj1.mensagem_log('\n====================================================================================================================================\n\nInício da execução')

            obj1.analisa_referencia()
            obj1.dados_completos_dos_fornecedores()
            obj1.listar_cnpjs()
            obj1.listar_cnpjs_exceções()

            obj1.mensagem_log_sem_horario('\nFornecedores analisados:')
            for emp in obj1.empresas:
                obj1.mensagem_log_sem_horario(f'{emp}')

            obj1.cria_diretorio()
            obj1.apaga_imagem()
            obj1.certidoes_n_encontradas()
            obj1.pdf_para_jpg()
            obj1.analisa_certidoes()

            obj1.mensagem_log_sem_horario('\nRESULTADO DA CONFERÊNCIA:')

            obj1.pega_cnpj()

            obj1.mensagem_log_sem_horario('\n\nCERTIDÕES QUE DEVEM SER ATUALIZADAS:\n')

            for emp in obj1.empresas_a_atualizar:
                obj1.mensagem_log_sem_horario(f'{emp} - {obj1.empresas_a_atualizar[emp][0:-1]} - CNPJ: {obj1.empresas_a_atualizar[emp][-1]}\n')

            obj1.apaga_imagem()

            tempo_final = time.time()
            tempo_de_execução = int((tempo_final - tempo_inicial))
            obj1.mensagem_log(
                f'\n\nTempo total de execução: {tempo_de_execução // 60} minutos e {tempo_de_execução % 60} segundos.')
            obj1.mensagem_log_sem_horario(
                '\n\n====================================================================================================================================\n')
            messagebox.showinfo('Analisou, miserávi!', 'Processo de análise de certidões executado com sucesso!')

    def selecionador_de_opções(self):
        print(self.variavel_de_opções.get())
        if self.variavel_de_opções.get() == 'Renomear arquivos':
            self.pdf_para_jpg_para_renomear_arquivo()
        elif self.variavel_de_opções.get() == 'Renomear todos os arquivos de uma pasta':
            self.pdf_para_jpg_renomear_conteudo_da_pasta()
        elif self.variavel_de_opções.get() == 'Renomear todas as certidões da lista de pagamento':
            self.renomeia()
        elif self.variavel_de_opções.get() == 'Selecione uma opção':
            messagebox.showwarning('Tem que escolher, fi!', 'Nenhuma opção selecionada!')

    def renomeia(self):
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        self.atualiza_urls()
        if not os.path.exists(self.urls[0][1]) or not os.path.exists(self.urls[1][1]) or not os.path.exists(self.urls[2][1])\
                or not os.path.exists(self.urls[4][1]) or not os.path.exists(self.urls[3][1]):
            self.checa_urls()
        else:
            obj1 = Certidao(dia, mes, ano)
            obj1.mensagem_log('\nProcesso de renomeação de certidões iniciado:\n')
            obj1.analisa_referencia()
            obj1.pega_fornecedores()
            obj1.apaga_imagem()
            obj1.pdf_para_jpg_renomear()
            obj1.gera_nome()
            obj1.apaga_imagem()
            messagebox.showinfo('Renomeou, miserávi!', 'Todas as certidões da listagem de pagamento foram renomeadas com sucesso!')

    def transfere_certidoes(self):
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        self.atualiza_urls()
        if not os.path.exists(self.urls[0][1]) or not os.path.exists(self.urls[1][1]) or not os.path.exists(
                self.urls[2][1]) \
                or not os.path.exists(self.urls[4][1]) or not os.path.exists(self.urls[3][1]):
            self.checa_urls()
        else:
            obj1 = Certidao(dia, mes, ano)
            obj1.analisa_referencia()
            obj1.pega_fornecedores()
            obj1.cria_certidoes_para_pagamento()

    def mescla_certidoes(self):
        dia = self.variavel.get()
        mes = self.variavel2.get()
        ano = self.variavel3.get()
        self.atualiza_urls()
        if not os.path.exists(self.urls[0][1]) or not os.path.exists(self.urls[1][1]) or not os.path.exists(
                self.urls[2][1]) \
                or not os.path.exists(self.urls[4][1]) or not os.path.exists(self.urls[3][1]):
            self.checa_urls()
        else:
            obj1 = Certidao(dia, mes, ano)
            obj1.analisa_referencia()
            obj1.pega_fornecedores()
            obj1.merge()

    def caminho_de_arquivo(self):
        self.arquivo_selecionado = filedialog.askopenfilenames(initialdir=f'{self.__file__()}/Certidões',
                                                               filetypes=(('Arquivos pdf','*.pdf'),("Todos os arquivos", '*.*')))
        numero_de_arquivos = 'Nenhum arquivo selecionado'
        if len(self.arquivo_selecionado) > 1:
            numero_de_arquivos = 'Multiplos arquivos selecionados'
        elif len(self.arquivo_selecionado) == 1:
            numero_de_arquivos = os.path.basename(self.arquivo_selecionado[0])
        self.caminho_do_arquivo = Label(self.frame_renomear, text=numero_de_arquivos, pady=0,
                                        padx=50, bg='white', fg='gray', font=('Helvetica', 9, 'bold'))
        self.caminho_do_arquivo.grid(row=0, column=2, padx=5, pady=0, ipadx=0, ipady=8, sticky=W+E)

    def pdf_para_jpg_para_renomear_arquivo(self):
        self.arquivo_selecionado = filedialog.askopenfilenames(initialdir=f'{self.__file__()}/Certidões',
                                                               filetypes=(
                                                               ('Arquivos pdf', '*.pdf'), ("Todos os arquivos", '*.*')))
        if self.arquivo_selecionado == 'Selecione os arquivos que deseja renomear' or list(self.arquivo_selecionado) == []:
            messagebox.showerror('Se não selecionar os arquivos, não vai rolar!', 'Selecione os arquivos que deseja renomear')
            print('Selecione os arquivos que deseja renomear')
        elif not os.path.exists(self.arquivo_selecionado[0]):
            print('O arquivo selecionado não existe.')
            messagebox.showerror('Esse arquivo é invenção da sua cabeça, parça!',
                                 'O arquivo selecionado não existe ou já foi renomeado!')
            self.caminho_do_arquivo = Label(self.frame_renomear, text='O arquivo selecionado não existe.', pady=0,
                                            padx=50, bg='white', fg='gray', font=('Helvetica', 9, 'bold'))
            self.caminho_do_arquivo.grid(row=0, column=2, padx=5, pady=0, ipadx=0, ipady=8, sticky=W + E)
        else:
            print('==================================================================================================\n'
                  'Criando imagem:\n')
            certidão_pdf = list(self.arquivo_selecionado)
            print(f'Arquivo que está sendo renomeado: {certidão_pdf}\n')
            for arquivo_a_renomear in certidão_pdf:
                os.chdir(arquivo_a_renomear[0:-((arquivo_a_renomear[::-1].find('/')+1))])
                pages = convert_from_path(arquivo_a_renomear, 300, last_page=1)
                certidão_convertida_para_jpg = f'{arquivo_a_renomear[:-4]}.jpg'
                pages[0].save(certidão_convertida_para_jpg, "JPEG")
                print('\nImagem criada com sucesso!\n')

                certidao_jpg = pytesseract.image_to_string(Image.open(certidão_convertida_para_jpg), lang='por')
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

                print('Renomeando certidão')
                for frase in padroes:
                    if frase in certidao_jpg:
                        if frase == 'GOVERNO DO DISTRITO FEDERAL':
                            try:
                                data = re.compile(datas2[frase])
                                procura = data.search(certidao_jpg)
                                datanome = procura.group()
                                separa = datanome.split('/')
                                junta = '-'.join(separa)
                            except AttributeError:
                                data = re.compile(datas[frase])
                                procura = data.search(certidao_jpg)
                                datanome = procura.group()
                                separa = datanome.split('/')
                                junta = '-'.join(separa)
                        else:
                            data = re.compile(datas[frase])
                            procura = data.search(certidao_jpg)
                            datanome = procura.group()
                            separa = datanome.split('/')
                            junta = '-'.join(separa)
                        if ':' in junta:
                            retira = junta.split(':')
                            volta = ' '.join(retira)
                            junta = volta
                        shutil.move(f'{certidão_convertida_para_jpg[0:-4]}.pdf', f'{valores[frase]} - {junta}.pdf')
                        os.unlink(certidão_convertida_para_jpg)
            print('\nProcesso de renomeação de certidão executado com sucesso!\n'
                  '===================================================================================================')
            messagebox.showinfo('Renomeou, miserávi!', 'Todas as certidões selecionadas foram renomeadas com sucesso!')



    def caminho_de_pastas(self):
        pasta = 'Nenhuma pasta selecionada'
        self.pasta_selecionada = filedialog.askdirectory(initialdir=f'{self.__file__()}/Certidões')
        if os.path.isdir(self.pasta_selecionada) and self.pasta_selecionada != f'{self.__file__()}/Certidões':
            pasta = self.pasta_selecionada
            self.caminho_da_pasta = Label(self.frame_renomear, text=os.path.basename(pasta), pady=0, padx=0, bg='white', fg='gray',
                       font=('Helvetica', 9, 'bold'))
            self.caminho_da_pasta.grid(row=1, column=2, columnspan=1, padx=5, pady=0, ipadx=0, ipady=8, sticky=W + E)
        else:
            self.caminho_da_pasta = Label(self.frame_renomear, text=pasta, pady=0, padx=0, bg='white',
                                          fg='gray',
                                          font=('Helvetica', 9, 'bold'))
            self.caminho_da_pasta.grid(row=1, column=2, columnspan=1, padx=5, pady=0, ipadx=0, ipady=8, sticky=W + E)


    def apaga_imagens_da_pasta(self):
            os.chdir(self.pasta_selecionada)
            for arquivo in os.listdir(self.pasta_selecionada):
                if arquivo.endswith(".jpg"):
                    os.unlink(f'{self.pasta_selecionada}/{arquivo}')

    def pdf_para_jpg_renomear_conteudo_da_pasta(self):
        self.pasta_selecionada = filedialog.askdirectory(initialdir=f'{self.__file__()}/Certidões')
        if self.pasta_selecionada == 'Selecione a pasta que deseja renomear' or self.pasta_selecionada =='':
            messagebox.showerror('Se não selecionar a pasta, não vai rolar!',
                                 'Selecione uma pasta que contenha certidões que precisam ser renomeadas.')
            print('nenhuma pasta selecionada')
            self.caminho_da_pasta = Label(self.frame_renomear, text='Nenhuma pasta selecionada', pady=0, padx=0, bg='white',
                                          fg='gray',
                                          font=('Helvetica', 9, 'bold'))
            self.caminho_da_pasta.grid(row=1, column=2, columnspan=1, padx=5, pady=0, ipadx=0, ipady=8, sticky=W + E)
        else:
            print('==================================================================================================\n'
                  'Criando imagens:\n')
            os.chdir(self.pasta_selecionada)
            self.apaga_imagens_da_pasta()
            for pdf_file in os.listdir(self.pasta_selecionada):
                if '00.MERGE' in pdf_file:
                    if not os.path.isdir(f'{self.pasta_selecionada}/Mesclados'):
                        os.makedirs(f'{self.pasta_selecionada}/Mesclados')
                        shutil.move(pdf_file, f'{self.pasta_selecionada}/Mesclados/{pdf_file}')
                    else:
                        shutil.move(pdf_file, f'{self.pasta_selecionada}/Mesclados/{pdf_file}')

                elif pdf_file.endswith(".pdf"):
                    print(pdf_file[:-4])
                    pages = convert_from_path(pdf_file, 300, last_page=1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")

            print(f'\nRenomeando certidões da pasta {self.pasta_selecionada}:\n\n')
            os.chdir(f'{self.pasta_selecionada}')
            origem = f'{self.pasta_selecionada}'
            for imagem in os.listdir(origem):
                if imagem.endswith(".jpg"):
                    certidao = pytesseract.image_to_string(Image.open(f'{origem}/{imagem}'), lang='por')
                    padroes = ['FGTS - CRF', 'Brasília,', 'JUSTIÇA DO TRABALHO', 'MINISTÉRIO DA FAZENDA',
                                   'GOVERNO DO DISTRITO FEDERAL']
                    valores = {'FGTS - CRF': 'FGTS', 'Brasília,': 'GDF', 'JUSTIÇA DO TRABALHO': 'TST',
                                   'MINISTÉRIO DA FAZENDA': 'UNIÃO', 'GOVERNO DO DISTRITO FEDERAL': 'GDF'}
                    datas = {'FGTS - CRF': 'a (\d\d)/(\d\d)/(\d\d\d\d)',
                                 'Brasília,': 'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                              '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)',
                                 'JUSTIÇA DO TRABALHO': 'Validade: (\d\d)/(\d\d)/(\d\d\d\d)',
                                 'MINISTÉRIO DA FAZENDA': 'Válida até (\d\d)/(\d\d)/(\d\d\d\d)',
                                 'GOVERNO DO DISTRITO FEDERAL': 'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                                                '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                                                                '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)'}
                    datas2 = {
                            'GOVERNO DO DISTRITO FEDERAL': 'Válida até (\d) de (janeiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?'
                                                           '(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)?(Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?'
                                                           '(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)'}
                    for frase in padroes:
                        if frase in certidao:
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
                            print(imagem.split()[0])
            self.apaga_imagens_da_pasta()
            print('\nProcesso de renomeação de certidões executado com sucesso!')
            messagebox.showinfo('Renomeou, miserávi!', 'Todas as certidões da pasta selecionada foram renomeadas com sucesso!')

tela = Tk()

objeto_tela = Analisador(tela)
tela.resizable(False, False)
tela.title('GEOF - Analisador de certidões')
#tela.iconbitmap('D:/Leiturapdf/GEOF_logo.ico')
tela.config(menu=objeto_tela.menu_certidões)

tela.mainloop()
