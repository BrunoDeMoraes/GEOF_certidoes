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
        self.caminho_xls = '//hrg-74977/GEOF/CERTIDÕES/Análise/atual.xlsx'
        self.wb = openpyxl.load_workbook(self.caminho_xls)
        self.pag = self.wb['PAGAMENTO']
        self.forn = self.wb['FORNECEDORES']
        self.listareferencia = []
        self.referencia = 0
        self.datapag = f'CERTIDÕES PARA {self.dia}/{self.mes}/{self.ano}'
        self.empresas = {}
        self.pdf_dir = '//hrg-74977/GEOF/CERTIDÕES/Certidões2'
        self.percentual = 0
        self.lista_de_cnpj = {}
        self.lista_de_cnpj_exceções = {}
        self.orgaos = ['UNIÃO', 'TST', 'FGTS', 'GDF']
        self.empresasdic = {}
        self.empresas_a_atualizar = {}
        self.caminho_de_log = f'//hrg-74977/GEOF/CERTIDÕES/Logs de conferência/{self.ano}-{self.mes}-{self.dia}.txt'
        self.pasta_de_trabalho = f'//hrg-74977/GEOF/HRG/PDPAS 2020/PAGAMENTO/{self.ano}-{self.mes}-{self.dia}'
        self.pagamento_por_data = f'//hrg-74977/GEOF/CERTIDÕES/Pagamentos/{self.ano}-{self.mes}-{self.dia}'

    def __file__(self):
        caminho_py = __file__
        caminho_do_dir = caminho_py.split('\\')
        caminho_uso = ('/').join(caminho_do_dir[0:-1])
        return caminho_uso

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
        for registro in self.lista_de_urls:
            print(registro)
            conexao.close()

    def pega_referencia(self):
        if os.path.exists(f'{self.pasta_de_trabalho}'):
            print('Pasta para inclusão de arquivos de pagamento localizada.')
        else:
            os.makedirs(f'{self.pasta_de_trabalho}')
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
            raise Exception(f'A data especificada foi encontrada nas células {self.listareferencia} da planilha de pagamentos: \\\hrg-74977\GEOF\CERTIDÕES\Análise\\atual.xlsx.'
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
                        self.lista_de_cnpj_exceções[celula.value] = ' '.join(nome_da_empresa[0:len(nome_da_empresa) - 1])
        return self.lista_de_cnpj_exceções

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
        if os.path.exists(f'{self.pagamento_por_data}'):
            print('Já existe pasta contendo certidões para pagamento na data informada.')
            messagebox.showwarning('FICA CALMO!!!', f'''Já existe pasta contendo certidões para pagamento na data informada!

Se deseja fazer nova transferência apague o diretório:
{self.pagamento_por_data}''')
        else:
            os.makedirs(self.pagamento_por_data)
            for emp in self.empresas:
                pasta_da_empresa = f'{self.pdf_dir}/{str(emp)}'
                os.makedirs(f'{self.pagamento_por_data}/{emp}')
                os.chdir(f'{pasta_da_empresa}')
                for pdf_file in os.listdir(f'{pasta_da_empresa}'):
                    if pdf_file.endswith(".pdf"):
                        shutil.copy(f'{pasta_da_empresa}/{pdf_file}', f'{self.pagamento_por_data}/{emp}/{pdf_file}')
            self.mensagem_log_sem_horario(f'As certidões referentes ao pagamento com data limite para a data de {self.dia}/{self.mes}/{self.ano} foram transferidas para respectiva pasta de pagamento.')
            messagebox.showinfo('Transferiu, miserávi!', 'As certidões que validam o pagamento foram transferidas com sucesso!')

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
            messagebox.showerror('Tá faltando coisa, mano!', f'''Algumas certidões não foram encontradas!
Consulte o arquivo de log, resolva as pendências indicadas e então execute novamente a análise.''')
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
                    raise Exception(f'''Não foi possível localizar o CNPJ da empresa {emp} na planilha FORNECEDORES do arquivo:
{self.caminho_xls}.
Verifique se há registro de CNPJ para a empresa ou se o nome informado na planilha PAGAMENTO é idêntico ao inserido na planilha FORNECEDORES.''')

                if len(self.empresas[emp]) > 3:
                    if val == True and cnpj_para_comparação == self.empresas[emp][3]:
                        empresadic[self.orgaos[index]] = 'OK-MATRIZ'
                    elif val == True and cnpj_para_comparação == self.empresas[emp][1]:
                        empresadic[self.orgaos[index]] = 'OK'
                    elif cnpj_para_comparação != self.empresas[emp][1] and cnpj_para_comparação != self.empresas[emp][3]:
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


    def merge(self):
        if os.path.exists(f'{self.pasta_de_trabalho}/Merge'):
            print('Já existe pasta para mesclagem na data informada')
            messagebox.showwarning('FICA CALMO!!!', f'''Já existe pasta para mesclagem na data informada!

Se deseja fazer nova mesclagem apague o diretório:
{self.pasta_de_trabalho}/Merge.''')
        else:
            os.makedirs(f'{self.pasta_de_trabalho}/Merge')
            os.chdir(self.pasta_de_trabalho)
            for arquivo_pdf in os.listdir(self.pasta_de_trabalho):
                os.chdir(self.pasta_de_trabalho)
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
                        if 'falha' not in validação_de_partes_do_nome:
                            print(emp)
                            pdf_temporário = PyPDF2.PdfFileWriter()
                            pagamento = open(arquivo_pdf, 'rb')
                            pagamento_lido = PyPDF2.PdfFileReader(pagamento, strict=False)
                            for página in range(pagamento_lido.numPages):
                                objeto_pagina = pagamento_lido.getPage(página)
                                pdf_temporário.addPage(objeto_pagina)
                            pasta_da_empresa = f'{self.pagamento_por_data}/{emp}'
                            os.chdir(pasta_da_empresa)
                            for arquivo_certidão in os.listdir(pasta_da_empresa):
                                if '00.MERGE' not in arquivo_certidão:
                                    certidão = open(arquivo_certidão, 'rb')
                                    certidão_lida = PyPDF2.PdfFileReader(certidão)
                                    for página_da_certidão in range(certidão_lida.numPages):
                                        objeto_pagina_da_certidão = certidão_lida.getPage(página_da_certidão)
                                        pdf_temporário.addPage(objeto_pagina_da_certidão)
                            compilado = open(f'{self.pasta_de_trabalho}/Merge/{arquivo_pdf[0:-4]}_mesclado.pdf','wb')
                            pdf_temporário.write(compilado)
                            compilado.close()
                            pagamento.close()
                            certidão.close()
            messagebox.showinfo('Mesclou, miserávi!!!', 'Digitalizações de pagamentos e respectivas certidões mescladas com sucesso!')


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

    def confere_data(self, certidao, emp):
        self.listar_cnpjs()
        self.listar_cnpjs_exceções()
        padrão_cnpj = re.compile('(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)')
        try:
            validação_de_cnpj = padrão_cnpj.search(certidao).group()
        except AttributeError:
            self.mensagem_log(f'Execução interrompida!!!\nNão foi possível encontrar o padrão de CNPJ na certidão UNIÃO da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
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
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à empresa {self.lista_de_cnpj[validação_de_cnpj]}\n')
        elif validação_de_cnpj in self.lista_de_cnpj_exceções:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à matriz da empresa {self.lista_de_cnpj_exceções[validação_de_cnpj]}\n')
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

    def confere_data(self, certidao, emp):
        self.listar_cnpjs()
        padrão_cnpj = re.compile('(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)')
        try:
            validação_de_cnpj = padrão_cnpj.search(certidao).group()
        except AttributeError:
            self.mensagem_log(f'Execução interrompida!!!\nNão foi possível encontrar o padrão de CNPJ na certidão TST da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
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

    def confere_data(self, certidao, emp):
        self.listar_cnpjs()
        padrão_cnpj = re.compile('(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)')
        try:
            validação_de_cnpj = padrão_cnpj.search(certidao).group()
        except AttributeError:
            self.mensagem_log(f'Execução interrompida!!!\nNão foi possível encontrar o padrão de CNPJ na certidão FGTS da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.')
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

    def confere_data(self, certidao, emp):
        self.listar_cnpjs()
        padrão_cnpj = re.compile('(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)')
        try:
            validação_de_cnpj = padrão_cnpj.search(certidao).group()
        except AttributeError:
            self.mensagem_log(f'''Execução interrompida!!!
Não foi possível encontrar o padrão de CNPJ na certidão GDF da empresa {emp}.
O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            messagebox.showerror('Esse arquivo não rola!', f'''Não foi possível encontrar o padrão de CNPJ na certidão GDF da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão GDF da empresa {emp} inválido.')
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


            padrao = re.compile('Válida até (\d)?(\d\d)? de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
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
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à empresa {self.lista_de_cnpj[validação_de_cnpj]}\n')
        else:
            self.mensagem_log_sem_horario(f'   O CNPJ encontrado, {validação_de_cnpj}, não possui correspondência\n')
        return (data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento), validação_de_cnpj