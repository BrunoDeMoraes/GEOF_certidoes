import os
import shutil
from tkinter import *
from tkinter import messagebox

import PyPDF2
import openpyxl
import pytesseract
from PIL import Image
from pdf2image import convert_from_path

from conexao import Conexao
from constantes import DATA_NAO_ENCONTRADA
from constantes import DATAS_MULTIPLAS
from constantes import ORGAOS
from constantes import PASTA_CRIADA
from constantes import PASTA_LOCALIZADA
from constantes import PLANILHAS
from constantes import REFERENCIA
from log import Log


class Certidao(Log, Conexao):
    def __init__(self, dia, mes, ano):
        self.dia = dia
        self.mes = mes
        self.ano = ano

        self.lista_de_urls = self.consulta_urls()

        self.caminho_xls = self.lista_de_urls[0][1]
        self.wb = openpyxl.load_workbook(self.caminho_xls)
        self.forn = self.wb[PLANILHAS[0]]
        self.pag = self.wb[PLANILHAS[1]]
        self.checagem_de_planilhas()

        self.listareferencia = []
        self.empresas = {}
        self.percentual = 0
        self.lista_de_cnpj = {}
        self.lista_de_cnpj_exceções = {}
        self.empresasdic = {}
        self.empresas_a_atualizar = {}

        self.pasta_de_certidões = self.lista_de_urls[1][1]
        self.caminho_de_log = (
            f'{self.lista_de_urls[2][1]}/{self.ano}-{self.mes}-{self.dia}.txt'
        )
        self.comprovantes_de_pagamento = (
            f'{self.lista_de_urls[3][1]}/{self.ano}-{self.mes}-{self.dia}'
        )
        self.certidões_para_pagamento = (
            f'{self.lista_de_urls[4][1]}/{self.ano}-{self.mes}-{self.dia}'
        )

    def checagem_de_planilhas(self):
        try:
            self.forn and self.pag
        except KeyError:
            messagebox.showerror(PLANILHAS[2], PLANILHAS[3])

    def checa_pasta_de_comprovantes(self):
        if os.path.exists(self.comprovantes_de_pagamento):
            print(PASTA_LOCALIZADA)
        else:
            os.makedirs(self.comprovantes_de_pagamento)
            print(PASTA_CRIADA)

    def pega_referencia(self):
        self.checa_pasta_de_comprovantes()
        data_para_pagamento = (
            f'CERTIDÕES PARA {self.dia}/{self.mes}/{self.ano}'
        )

        for linha in self.pag['A1':'F500']:
            for celula in linha:
                if celula.value != data_para_pagamento:
                    continue
                elif (
                        celula.value == data_para_pagamento
                        and celula.coordinate not in self.listareferencia
                ):
                    self.listareferencia.append(celula.coordinate)
        return self.listareferencia

    def analisa_referencia(self):
        self.pega_referencia()
        if len(self.listareferencia) == 0:
            self.mensagem_de_log_completa(
                DATA_NAO_ENCONTRADA[1],
                self.caminho_de_log
            )
            messagebox.showerror(
                DATA_NAO_ENCONTRADA[0],
                DATA_NAO_ENCONTRADA[1]
            )
            raise Exception(DATA_NAO_ENCONTRADA[1])

        elif len(self.listareferencia) > 1:
            mensagem_de_erro = (
                    f'{DATAS_MULTIPLAS[1]}\n\n'
                    f'{self.listareferencia}\n'
                    f'{DATAS_MULTIPLAS[2]}'
                 )
            self.mensagem_de_log_completa(
                mensagem_de_erro,
                self.caminho_de_log
            )
            messagebox.showerror(
                DATAS_MULTIPLAS[0],
                mensagem_de_erro
            )
            raise Exception(mensagem_de_erro)
        else:
            self.mensagem_de_log_completa(
                f'\n{REFERENCIA}{self.listareferencia[0]}',
                self.caminho_de_log
            )

#continuar refatoração desse ponto
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
        self.mensagem_de_log_completa(f'\nNúmero de novas pastas criadas: {len(novos_dir)} - {novos_dir}.', self.caminho_de_log)

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
            self.mensagem_de_log_simples(
                f'As certidões referentes ao pagamento com data limite para a data de {self.dia}/{self.mes}/{self.ano} foram transferidas para respectiva pasta de pagamento.', self.caminho_de_log)
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
            for orgao in ORGAOS:
                if orgao not in itens:
                    faltando.append(orgao)
            if faltando != []:
                try:
                    self.empresas[emp][2]
                except:
                    messagebox.showerror('Problema com o xlsx', 'O arquivo fonte de dados XLSX parece não ter sido atualizado corretamente.\n\n'
                                                                                'Tente atualizar a planilha FORNECEDORES usando a oção de colagem  que insere apenas "Valores"')
                self.mensagem_de_log_completa(f'Para a empresa {emp} não foram encontradas as certidões {faltando} - CNPJ: {self.empresas[emp][2]}', self.caminho_de_log)
                total_faltando += 1
        if total_faltando != 0:
            self.mensagem_de_log_completa(f'Adicione as certidões às respectivas pastas informadas e execute novamente o programa.', self.caminho_de_log)
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
                if pdf_file.endswith(".pdf") and pdf_file.split()[0] in ORGAOS:
                    pages = convert_from_path(pdf_file, 300, last_page=1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")

    def analisa_certidoes(self, lista_de_objetos):
        #objUniao = Uniao(self.dia, self.mes, self.ano)
        #objTst = Tst(self.dia, self.mes, self.ano)
        #objFgts = Fgts(self.dia, self.mes, self.ano)
        #objGdf = Gdf(self.dia, self.mes, self.ano)
        lista_objetos = lista_de_objetos
        self.mensagem_de_log_completa('\nInicio da conferência de datas de emissão e vencimento:', self.caminho_de_log)
        print(f'Total executado: {self.percentual}%')

        for emp in self.empresas:
            empresadic = {}
            index = 0
            self.mensagem_de_log_completa(f'\n{emp}', self.caminho_de_log)
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
                    self.mensagem_de_log_completa(f'Não foi possível localizar o CNPJ da '
                                      f'empresa {emp} na planilha FORNECEDORES'
                                      f' do arquivo: {self.caminho_xls}.\n\n'
                                      f'Verifique se há registro de CNPJ para a'
                                      f' empresa ou se o nome informado na'
                                      f' planilha PAGAMENTO é idêntico ao '
                                      f'inserido na planilha FORNECEDORES.', self.caminho_de_log)
                    raise Exception(
                        f'''Não foi possível localizar o CNPJ da empresa {emp} na planilha FORNECEDORES do arquivo:
{self.caminho_xls}.
Verifique se há registro de CNPJ para a empresa ou se o nome informado na planilha PAGAMENTO é idêntico ao inserido na planilha FORNECEDORES.''')

                if len(self.empresas[emp]) > 3:
                    if val == True and cnpj_para_comparação == self.empresas[emp][3]:
                        empresadic[ORGAOS[index]] = 'OK-MATRIZ'
                    elif val == True and cnpj_para_comparação == self.empresas[emp][1]:
                        empresadic[ORGAOS[index]] = 'OK'
                    elif cnpj_para_comparação != self.empresas[emp][1] and cnpj_para_comparação != self.empresas[emp][
                        3]:
                        empresadic[ORGAOS[index]] = 'CNPJ-ERRO'
                    else:
                        empresadic[ORGAOS[index]] = 'INCOMPATÍVEL'
                else:
                    if val == True and cnpj_para_comparação == self.empresas[emp][1]:
                        empresadic[ORGAOS[index]] = 'OK'
                    elif cnpj_para_comparação != self.empresas[emp][1]:
                        empresadic[ORGAOS[index]] = 'CNPJ-ERRO'
                    else:
                        empresadic[ORGAOS[index]] = 'INCOMPATÍVEL'
                index += 1
            self.empresasdic[emp] = empresadic

    def atualizar(self):
        numerador = 0
        for emp in self.empresasdic:
            self.mensagem_de_log_simples(f'{numerador + 1 :>2} - {emp}\n{self.empresasdic[emp]}\n', self.caminho_de_log)
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
        self.mensagem_de_log_completa(
            '\nImagens criadas com sucesso!',
            self.caminho_de_log
        )
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