import os
import shutil
import re
from datetime import date
from tkinter import *
from tkinter import messagebox

import PyPDF2
import openpyxl
import pytesseract
import PIL.Image
from pdf2image import convert_from_path

from barra_de_progresso import Barra
from conexao import Conexao
from constantes import ANALISADOS
from constantes import ARQUIVO_CORROMPIDO
from constantes import ATUALIZAR_XLSX
from constantes import CERTIDOES_FALTANDO
from constantes import CERTIDOES_TRANSFERIDAS
from constantes import CNPJ_VAZIO
from constantes import CONFERENCIA
from constantes import CRIANDO_IMAGENS
from constantes import DADOS_DO_FORNECEDOR_COM_ERRO
from constantes import DATAS_MULTIPLAS
from constantes import DATA_NAO_ENCONTRADA
from constantes import DIGITALIZADOS_MESCLADOS
from constantes import IDENTIFICADOR_DE_CERTIDAO
from constantes import IDENTIFICADOR_DE_VALIDADE, IDENTIFICADOR_DE_VALIDADE_2
from constantes import IDENTIFICADOR_TRADUZIDO
from constantes import INICIO_DA_ANALISE
from constantes import INICIO_DA_EXECUCAO
from constantes import LINHA_FINAL
from constantes import ORGAOS
from constantes import PADRAO_CNPJ
from constantes import PASTA_CRIADA
from constantes import PASTA_DE_MESCLAGEM_EXISTENTE
from constantes import PASTA_DE_PAGAMENTO
from constantes import PASTA_LOCALIZADA
from constantes import PASTA_NAO_ENCONTRADA
from constantes import PENDENCIAS
from constantes import PLANILHAS
from constantes import REFERENCIA
from constantes import RENOMEACAO_EXECUTADA
from log import Log


class Certidao(Log, Conexao, Barra):
    def __init__(self, dia, mes, ano):
        self.dia_c = dia
        self.mes_c = mes
        self.ano_c = ano

        self.lista_de_urls = self.consulta_urls()

        self.listareferencia = []
        self.empresas = {}
        self.percentual = 0
        self.lista_de_cnpj = {}
        self.lista_de_cnpj_exceções = {}
        self.empresasdic = {}
        self.empresas_a_atualizar = {}

        self.pasta_de_certidões = self.lista_de_urls[1][1]
        self.caminho_de_log = (
            f'{self.lista_de_urls[2][1]}/{self.ano_c}-{self.mes_c}-'
            f'{self.dia_c}.txt'
        )
        self.comprovantes_de_pagamento = (
            f'{self.lista_de_urls[3][1]}/{self.ano_c}-{self.mes_c}-'
            f'{self.dia_c}'
        )
        self.certidões_para_pagamento = (
            f'{self.lista_de_urls[4][1]}/{self.ano_c}-{self.mes_c}-'
            f'{self.dia_c}'
        )
        self.caminho_xls = self.lista_de_urls[0][1]
        self.wb = openpyxl.load_workbook(self.caminho_xls)
        self.checagem_de_planilhas()
        self.forn = self.wb[PLANILHAS[0]]
        self.pag = self.wb[PLANILHAS[1]]

    def checagem_de_planilhas(self):
        if (PLANILHAS[0] or PLANILHAS[1]) not in self.wb.sheetnames:
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
            f'CERTIDÕES PARA {self.dia_c}/{self.mes_c}/{self.ano_c}'
        )

        for linha in self.pag['A1':'F1000']:
            for celula in linha:
                if celula.value != data_para_pagamento:
                    continue
                elif (
                        celula.value == data_para_pagamento
                        and celula.coordinate not in self.listareferencia
                ):
                    self.listareferencia.append(celula.coordinate)
        return self.listareferencia

    def referencia_vazia(self):
        self.mensagem_de_log_completa(
            DATA_NAO_ENCONTRADA[1],
            self.caminho_de_log
        )
        messagebox.showerror(
            DATA_NAO_ENCONTRADA[0],
            DATA_NAO_ENCONTRADA[1]
        )
        raise Exception(DATA_NAO_ENCONTRADA[1])

    def referencia_multipla(self):
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

    def analisa_referencia(self):
        self.pega_referencia()
        if len(self.listareferencia) == 0:
            self.referencia_vazia()
        elif len(self.listareferencia) > 1:
            self.referencia_multipla()
        else:
            self.mensagem_de_log_completa(
                f'\n{REFERENCIA}{self.listareferencia[0]}',
                self.caminho_de_log
            )

    def pega_fornecedores(self):
        referencia = self.listareferencia[0]
        desloca = 2
        coluna = referencia[0]
        linha = int(referencia[1:])
        while self.pag[coluna + str(linha + desloca)].value != None:
            empresa = self.pag[coluna + str(linha + desloca)].value.split()
            if len(empresa) > 2:
                self.empresas[' '.join(empresa[0:len(empresa) - 1])] = [
                    ' '.join(empresa)
                ]
            else:
                self.empresas[empresa[0]] = [' '.join(empresa)]
            desloca += 1
        return self.empresas

    def inclui_cnpj_em_fornecedores(self):
        for emp in self.empresas:
            for linha in self.forn['A6':'A1000']:
                for celula in linha:
                    if celula.value != self.empresas[emp][0]:
                        continue
                    else:
                        padrão_cnpj = re.compile(PADRAO_CNPJ)
                        celula_de_cnpj = self.forn[f'F{celula.row}'].value
                        if (
                                celula_de_cnpj == None
                                or padrão_cnpj.search(celula_de_cnpj) == None
                        ):
                            messagebox.showerror(
                                CNPJ_VAZIO[0], f'{emp} - {CNPJ_VAZIO[1]}'
                            )
                            raise Exception(f'{emp} - {CNPJ_VAZIO[1]}')
                        else:
                            self.empresas[emp].append(
                                str(self.forn[f'F{celula.row}'].value)
                            )
                            cnpj_formatado = self.empresas[emp][1]
                            cnpj_tratado = ''
                            for digito in cnpj_formatado:
                                if digito in '0123456789':
                                    cnpj_tratado += digito
                            self.empresas[emp].append(cnpj_tratado)
                            if self.forn[f'M{celula.row}'].value == None:
                                continue
                            else:
                                self.empresas[emp].append(
                                    self.forn[f'M{celula.row}'].value
                                )
                                cnpj_matriz_formatado = self.empresas[emp][3]
                                cnpj_matriz_tratado = ''
                                for digito in cnpj_matriz_formatado:
                                    if digito in '0123456789':
                                        cnpj_matriz_tratado += digito
                                self.empresas[emp].append(cnpj_matriz_tratado)
        return self.empresas

    def analisa_lista_de_fornecedores(self):
        self.mensagem_de_log_completa(
                INICIO_DA_EXECUCAO, self.caminho_de_log
        )
        self.analisa_referencia()
        self.pega_fornecedores()
        self.inclui_cnpj_em_fornecedores()
        self.listar_cnpjs()
        self.lista_de_cnpj_exceções

        self.mensagem_de_log_simples(ANALISADOS, self.caminho_de_log)

    def analisa_requisitos(self):
        for emp in self.empresas:
            self.mensagem_de_log_simples(f'{emp}', self.caminho_de_log)

        self.cria_diretorio()
        self.apaga_imagem()
        self.certidoes_n_encontradas()
        self.pdf_para_jpg_analise()
        self.destruir_barra_de_progresso()

    def listar_cnpjs(self):
        for emp in self.empresas:
            for linha in self.forn['F6':'F1000']:
                for celula in linha:
                    if celula.value == None:
                        continue
                    else:
                        nome_da_empresa = (
                            self.forn[f'A{celula.row}'].value.split()
                        )
                        self.lista_de_cnpj[celula.value] = ' '.join(
                            nome_da_empresa[0:len(nome_da_empresa) - 1]
                        )
        return self.lista_de_cnpj

    def listar_cnpjs_exceções(self):
        for emp in self.empresas:
            for linha in self.forn['M6':'M1000']:
                for celula in linha:
                    if celula.value == None:
                        continue
                    else:
                        nome_da_empresa = (
                            self.forn[f'A{celula.row}'].value.split()
                        )
                        self.lista_de_cnpj_exceções[celula.value] = ' '.join(
                            nome_da_empresa[0:len(nome_da_empresa) - 1])
        return self.lista_de_cnpj_exceções

    def cria_diretorio(self):
        novos_dir = []
        for emp in self.empresas:
            if os.path.isdir(f'{self.pasta_de_certidões}/{str(emp)}'):
                continue
            else:
                os.makedirs(f'{self.pasta_de_certidões}/{str(emp)}')
                novos_dir.append(emp)
        self.mensagem_de_log_completa(
            (f'\n{PASTA_CRIADA[1]}{len(novos_dir)} - {novos_dir}.'),
            self.caminho_de_log
        )

    def mensagens_da_função_cria_certidoes(self):
        self.mensagem_de_log_simples(
            (
                f'{CERTIDOES_TRANSFERIDAS[1]}{self.dia_c}/{self.mes_c}/'
                f'{self.ano_c}'
            ),
            self.caminho_de_log
        )
        messagebox.showinfo(
            CERTIDOES_TRANSFERIDAS[0],
            CERTIDOES_TRANSFERIDAS[1]
        )

    def cria_certidoes_para_pagamento(self):
        if os.path.exists(f'{self.certidões_para_pagamento}'):
            print(PASTA_DE_PAGAMENTO)
            messagebox.showwarning(
                PASTA_DE_PAGAMENTO[0],
                f'{PASTA_DE_PAGAMENTO[1]}{self.certidões_para_pagamento}')
        else:
            os.makedirs(self.certidões_para_pagamento)
            for emp in self.empresas:
                pasta_da_empresa = f'{self.pasta_de_certidões}/{str(emp)}'
                os.makedirs(f'{self.certidões_para_pagamento}/{emp}')
                os.chdir(f'{pasta_da_empresa}')
                for pdf_file in os.listdir(f'{pasta_da_empresa}'):
                    if pdf_file.endswith(".pdf"):
                        shutil.copy(
                            f'{pasta_da_empresa}/{pdf_file}',
                            (f'{self.certidões_para_pagamento}/{emp}/'
                             f'{pdf_file}')
                        )
            self.mensagens_da_função_cria_certidoes()

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
                    messagebox.showerror(
                        ATUALIZAR_XLSX[0],
                        ATUALIZAR_XLSX[1]
                    )
                self.mensagem_de_log_completa(
                    (
                        f'{emp}, CNPJ: {self.empresas[emp][2]}'
                        f'{CERTIDOES_FALTANDO[0]}{faltando} '
                    ),
                    self.caminho_de_log)
                total_faltando += 1
        if total_faltando != 0:
            self.mensagem_de_log_completa(
                CERTIDOES_FALTANDO[1],
                self.caminho_de_log
            )
            messagebox.showerror(CERTIDOES_FALTANDO[2], CERTIDOES_FALTANDO[3])
            raise Exception(CERTIDOES_FALTANDO[1])
        else:
            self.imagens_criadas = 0
            self.thread_barra_de_progresso(
                'Criando imagens', self.imagens_criadas
            )

    def pdf_para_jpg_analise(self):
        for emp in self.empresas:
            self.imagens_criadas += (1/len(self.empresas))*100
            os.chdir(f'{self.pasta_de_certidões}/{str(emp)}')
            for pdf_file in os.listdir(
                    f'{self.pasta_de_certidões}/{str(emp)}'
            ):
                if (
                        pdf_file.endswith(".pdf")
                        and pdf_file.split()[0] in ORGAOS
                ):

                    pages = convert_from_path(pdf_file, 300, last_page=1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")
                    self.valor_da_barra(self.imagens_criadas)

    def analisa_certidoes(self, lista_de_objetos):
        self.thread_barra_de_progresso(
            'Analisando certidões', self.percentual
        )
        self.mensagem_de_log_completa(INICIO_DA_ANALISE, self.caminho_de_log)
        print(f'Total executado: {self.percentual:.2f}%')

        for emp in self.empresas:
            empresadic = {}
            index = 0
            self.mensagem_de_log_completa(f'\n{emp}', self.caminho_de_log)
            for objeto in lista_de_objetos:
                objeto.empresas = self.empresas
                objeto.lista_de_cnpj = self.lista_de_cnpj
                certidao = objeto.pega_string(emp)
                self.percentual += (25 / len(self.empresas))
                print(f'   Total executado: {self.percentual:.2f}%')
                self.valor_da_barra(self.percentual)
                try:
                    val, cnpj_para_comparação = objeto.confere_data(
                        certidao, emp
                    )
                except:
                    self.destruir_barra_de_progresso()
                    raise
                try:
                    self.empresas[emp][1]
                except IndexError:
                    messagebox.showerror(
                        DADOS_DO_FORNECEDOR_COM_ERRO[0],
                        (
                            f'empresa: {emp}\n\n'
                            f'{DADOS_DO_FORNECEDOR_COM_ERRO[1]}'
                        )
                    )
                    self.mensagem_de_log_completa(
                        (
                            f'empresa: {emp}\n\n'
                            f'{DADOS_DO_FORNECEDOR_COM_ERRO[1]}'
                        ),
                        self.caminho_de_log
                    )
                    self.destruir_barra_de_progresso()
                    raise Exception(
                        f'empresa: {emp}\n\n{DADOS_DO_FORNECEDOR_COM_ERRO[1]}'
                    )

                print(f' Não sei o que é empresa[emp] - {self.empresas[emp]}')
                if len(self.empresas[emp]) > 3:
                    if val and cnpj_para_comparação == self.empresas[emp][3]:
                        empresadic[ORGAOS[index]] = 'OK-MATRIZ'
                    elif (
                            val
                            and cnpj_para_comparação == self.empresas[emp][1]
                    ):
                        empresadic[ORGAOS[index]] = 'OK'
                    elif (
                            cnpj_para_comparação != self.empresas[emp][1]
                            and cnpj_para_comparação != self.empresas[emp][3]
                    ):
                        empresadic[ORGAOS[index]] = 'CNPJ-ERRO'
                    else:
                        empresadic[ORGAOS[index]] = 'INCOMPATÍVEL'
                else:
                    if val and cnpj_para_comparação == self.empresas[emp][1]:
                        empresadic[ORGAOS[index]] = 'OK'
                    elif cnpj_para_comparação != self.empresas[emp][1]:
                        empresadic[ORGAOS[index]] = 'CNPJ-ERRO'
                    else:
                        empresadic[ORGAOS[index]] = 'INCOMPATÍVEL'
                index += 1
            self.empresasdic[emp] = empresadic

    def resultado_da_analise(self):
        self.mensagem_de_log_simples(CONFERENCIA, self.caminho_de_log)

        self.pega_cnpj()

        self.mensagem_de_log_simples(PENDENCIAS, self.caminho_de_log)

        for emp in self.empresas_a_atualizar:
            self.mensagem_de_log_simples(
                (f'{emp} - {self.empresas_a_atualizar[emp][0:-1]} '
                 f'- CNPJ: {self.empresas_a_atualizar[emp][-1]}\n'),
                self.caminho_de_log
            )

        self.apaga_imagem()
        self.destruir_barra_de_progresso()

    def atualizar(self):
        numerador = 0
        for emp in self.empresasdic:
            self.mensagem_de_log_simples(
                f'{numerador + 1 :>2} - {emp}\n{self.empresasdic[emp]}\n',
                self.caminho_de_log
            )
            numerador += 1
        for emp in self.empresasdic:
            certidoes_a_atualizar = []
            for orgao in self.empresasdic[emp]:
                if (
                        self.empresasdic[emp][orgao] == 'INCOMPATÍVEL'
                        or self.empresasdic[emp][orgao] == 'CNPJ-ERRO'
                ):
                    certidoes_a_atualizar.append(orgao)
            if len(certidoes_a_atualizar) > 0:
                self.empresas_a_atualizar[emp] = certidoes_a_atualizar

    def pega_cnpj(self):
        self.atualizar()
        for emp in self.empresas_a_atualizar:
            for linha in self.forn['A6':'A1000']:
                for celula in linha:
                    if celula.value == None:
                        continue
                    else:
                        empresa = celula.value.split()
                        nome_da_empresa = ''
                        if len(empresa) > 2:
                            nome_da_empresa = (
                                ' '.join(empresa[0:len(empresa) - 1])
                            )
                        else:
                            nome_da_empresa = empresa[0]
                        if nome_da_empresa != emp:
                            continue
                        else:
                            cnpj_formatado = (
                                str(self.forn['F' + str(celula.row)].value)
                            )
                            cnpj_tratado = ''
                            for digito in cnpj_formatado:
                                if digito in '0123456789':
                                    cnpj_tratado += digito
                            self.empresas_a_atualizar[emp].append(
                                cnpj_tratado
                            )

    def opcao_de_identificador_1(self, frase, documento):
        try:
            validade_separada = self.buscar_data(frase, documento, IDENTIFICADOR_DE_VALIDADE_2)
            validade = '-'.join(validade_separada)
        except AttributeError:
            validade_separada = self.buscar_data(frase, documento, IDENTIFICADOR_DE_VALIDADE)
            validade = '-'.join(validade_separada)
        return validade

    def opcao_de_identificador_2(self, frase, documento):
        if frase == 'CONTROLADORIA-GERAL DA UNIÃO':
            if 'Data consulta' in documento:
                print('Tá lá')

        validade_separada = self.buscar_data(frase, documento, IDENTIFICADOR_DE_VALIDADE)
        hoje = date.today()
        if validade_separada[-1] == '2028' and int(hoje.year) not in [2027, 2028]:
            validade_separada[-1] = '2023'
        validade = '-'.join(data_separada)
        return validade


    def buscar_data(self, frase, documento, id_validade):
        data = re.compile(
            id_validade[frase]
        )
        procura = data.search(documento)
        datanome = procura.group()
        separa = datanome.split('/')
        return separa


    def pdf_para_jpg_para_renomear_arquivos(self, arquivo_selecionado):
        print(CRIANDO_IMAGENS[0])

        self.renomeadas = 0
        self.thread_barra_de_progresso(
            'Renomeando certidões', self.renomeadas
        )

        try:
            certidão_pdf = list(arquivo_selecionado)

            for arquivo_a_renomear in certidão_pdf:
                ultima_barra: int = arquivo_a_renomear[::-1].find('/') + 1
                os.chdir(arquivo_a_renomear[0:-(ultima_barra)])

                pages = convert_from_path(
                    arquivo_a_renomear, 300, last_page=1
                )

                imagem_da_certidao = f'{arquivo_a_renomear[:-4]}.jpg'
                pages[0].save(imagem_da_certidao, "JPEG")
                certidao_jpg = pytesseract.image_to_string(
                    PIL.Image.open(imagem_da_certidao), lang='por'
                )
                self.renomeadas += (1 / len(certidão_pdf)) * 100
                self.valor_da_barra(self.renomeadas)

                for frase in IDENTIFICADOR_DE_CERTIDAO:
                    if frase in certidao_jpg:
                        if frase == 'GOVERNO DO DISTRITO FEDERAL':
                            junta = self.opcao_de_identificador_1(frase, certidao_jpg)
                        else:
                            junta = self.opcao_de_identificador_2(frase, certidao_jpg)
                        if ':' in junta:
                            retira = junta.split(':')
                            volta = ' '.join(retira)
                            junta = volta
                        print(f'Resultado {IDENTIFICADOR_TRADUZIDO[frase]} - {junta}.pdf')
                        shutil.move(
                            f'{imagem_da_certidao[0:-4]}.pdf',
                            f'{IDENTIFICADOR_TRADUZIDO[frase]} - {junta}.pdf'
                        )
                        os.unlink(imagem_da_certidao)

            print(RENOMEACAO_EXECUTADA[2])
            print(LINHA_FINAL)

            messagebox.showinfo(
                RENOMEACAO_EXECUTADA[0],
                RENOMEACAO_EXECUTADA[2]
            )
        finally:
            self.destruir_barra_de_progresso()

    def apaga_imagens_da_pasta(self):
            os.chdir(self.pasta_selecionada)
            for arquivo in os.listdir(self.pasta_selecionada):
                if arquivo.endswith(".jpg"):
                    os.unlink(f'{self.pasta_selecionada}/{arquivo}')

    def pdf_para_jpg_renomear_conteudo_da_pasta(self):
        self.renomeadas = 0
        self.thread_barra_de_progresso(
            'Renomeando certidões', self.renomeadas
        )
        try:
            print(CRIANDO_IMAGENS[0])
            os.chdir(self.pasta_selecionada)

            for pdf_file in os.listdir(self.pasta_selecionada):

                if pdf_file.endswith(".pdf"):
                    print(pdf_file[:-4])
                    pages = convert_from_path(pdf_file, 300, last_page=1)
                    pdf_file = pdf_file[:-4]
                    pages[0].save(f"{pdf_file}.jpg", "JPEG")

            print(
                f'\nRenomeando certidões da pasta {self.pasta_selecionada}:\n'
                f'\n'
            )

            os.chdir(f'{self.pasta_selecionada}')
            origem = f'{self.pasta_selecionada}'

            for imagem in os.listdir(origem):

                self.renomeadas += (1 / len(os.listdir(origem))) * 100
                self.valor_da_barra(self.renomeadas)

                if imagem.endswith(".jpg"):
                    certidao = pytesseract.image_to_string(
                        PIL.Image.open(f'{origem}/{imagem}'), lang='por'
                    )

                    for frase in IDENTIFICADOR_DE_CERTIDAO:
                        if frase in certidao:
                            if frase == IDENTIFICADOR_DE_CERTIDAO[4]:
                                junta = self.opcao_de_identificador_1(frase, certidao)
                            else:
                                junta = self.opcao_de_identificador_2(frase, certidao)
                            if ':' in junta:
                                retira = junta.split(':')
                                volta = ' '.join(retira)
                                junta = volta
                            shutil.move(
                                f'{origem}/{imagem[0:-4]}.pdf',
                                f'{IDENTIFICADOR_TRADUZIDO[frase]} - '
                                f'{junta}.pdf'
                            )
                            print(imagem.split()[0])

            self.apaga_imagens_da_pasta()
            print(RENOMEACAO_EXECUTADA[2])

            messagebox.showinfo(
                RENOMEACAO_EXECUTADA[0],
                RENOMEACAO_EXECUTADA[3]
            )
        finally:
            self.destruir_barra_de_progresso()

    def pdf_para_jpg_renomeacao_geral(self):
        self.imagens_criadas = 0
        self.thread_barra_de_progresso(
            'Criando imagens', self.imagens_criadas
        )
        try:
            print(CRIANDO_IMAGENS[0])
            for emp in self.empresas:
                self.imagens_criadas += (1 / len(self.empresas)) * 100
                os.chdir(f'{self.pasta_de_certidões}/{str(emp)}')
                for pdf_file in os.listdir(
                        f'{self.pasta_de_certidões}/{str(emp)}'
                ):
                    if pdf_file.endswith(".pdf"):
                        pages = convert_from_path(pdf_file, 300, last_page=1)
                        pdf_file = pdf_file[:-4]
                        pages[0].save(f"{pdf_file}.jpg", "JPEG")
                        self.percentual += (25 / len(self.empresas))
                        print(
                            f'Total de imagens criadas: '
                            f'{self.percentual:.2f}%'
                        )
                        self.valor_da_barra(self.percentual)
            self.mensagem_de_log_completa(
                CRIANDO_IMAGENS[1],
                self.caminho_de_log
            )
            self.imagens_criadas = 0
        finally:
            self.destruir_barra_de_progresso()

    def gera_nome(self):
        self.percentual = 0
        self.renomeadas = 0
        self.thread_barra_de_progresso(
            'Renomeando certidões', self.renomeadas
        )
        try:
            print('\nRenomeando certidões:\n\n')
            for emp in self.empresas:
                self.renomeadas += (1 / len(self.empresas)) * 100
                os.chdir(f'{self.pasta_de_certidões}/{(emp)}')
                origem = f'{self.pasta_de_certidões}/{emp}'
                for imagem in os.listdir(origem):
                    if imagem.endswith(".jpg"):
                        certidao = pytesseract.image_to_string(
                            PIL.Image.open(f'{origem}/{imagem}'),
                            lang='por'
                        )
                        for frase in IDENTIFICADOR_DE_CERTIDAO:
                            if frase in certidao:
                                self.percentual += (25 / len(self.empresas))
                                print(
                                    (
                                        f'{emp} - certidão '
                                        f'{IDENTIFICADOR_TRADUZIDO[frase]} '
                                        f'renomeada - Total executado: '
                                        f'{self.percentual:.2f}%\n'
                                    )
                                )
                                if frase == 'GOVERNO DO DISTRITO FEDERAL':
                                    junta = self.opcao_de_identificador_1(
                                        frase, certidao
                                    )
                                else:
                                    junta = self.opcao_de_identificador_2(
                                        frase, certidao
                                    )
                                if ':' in junta:
                                    retira = junta.split(':')
                                    volta = ' '.join(retira)
                                    junta = volta
                                shutil.move(
                                    f'{origem}/{imagem[0:-4]}.pdf',
                                    (
                                        f'{IDENTIFICADOR_TRADUZIDO[frase]} - '
                                        f'{junta}.pdf'
                                    )
                                )
                self.valor_da_barra(self.renomeadas)
            print(RENOMEACAO_EXECUTADA[2])
        finally:
            self.destruir_barra_de_progresso()

    def rotina_de_renomeacao_geral(self):
        self.analisa_referencia()
        self.pega_fornecedores()
        self.apaga_imagem()
        self.pdf_para_jpg_renomeacao_geral()
        self.gera_nome()
        self.apaga_imagem()

        messagebox.showinfo(
            RENOMEACAO_EXECUTADA[0],
            RENOMEACAO_EXECUTADA[1]
        )

    def rotina_de_transferencia_de_certidoes(self):
        self.analisa_referencia()
        self.pega_fornecedores()
        self.cria_certidoes_para_pagamento()

    def merge(self):
        nomes_errados = []
        if os.path.exists(f'{self.comprovantes_de_pagamento}/Mesclados'):
            print(PASTA_DE_MESCLAGEM_EXISTENTE[1])
            messagebox.showwarning(
                PASTA_DE_MESCLAGEM_EXISTENTE[0],
                f'{PASTA_DE_MESCLAGEM_EXISTENTE[1]}'
                f'{self.comprovantes_de_pagamento}/Mesclados.')
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
                        retira_espaço_do_arquivo = arquivo_pdf.replace(
                            ' ', '-'
                        )
                        arquivo_separado = retira_espaço_do_arquivo.split('-')
                        while len(arquivo_separado) <= len(nome_separado):
                            arquivo_separado.append('item de correção')
                        contador = 0
                        for parte_do_nome in nome_separado:

                            if (
                                    nome_separado[contador]
                                    == arquivo_separado[contador + 1]
                            ):
                                validação_de_partes_do_nome.append('OK')
                                contador += 1
                            else:
                                validação_de_partes_do_nome.append('falha')
                                contador += 1
                        if 'falha' not in validação_de_partes_do_nome:
                            print(emp)
                            pdf_temporário = PyPDF2.PdfWriter()
                            pagamento = open(arquivo_pdf, 'rb')
                            try:
                                pagamento_lido = PyPDF2.PdfReader(
                                    pagamento, strict=False
                                )
                            except:
                                messagebox.showerror(
                                    ARQUIVO_CORROMPIDO[0],
                                    f'{ARQUIVO_CORROMPIDO[1]}{arquivo_pdf}'
                                )
                                self.destruir_barra_de_progresso()
                            for página in range(len(pagamento_lido.pages)):
                                objeto_pagina = pagamento_lido.pages[página]
                                pdf_temporário.add_page(objeto_pagina)
                            pasta_da_empresa = (
                                f'{self.certidões_para_pagamento}/{emp}'
                            )
                            os.chdir(pasta_da_empresa)
                            for arquivo_certidão in os.listdir(
                                    pasta_da_empresa
                            ):
                                if '00.MERGE' not in arquivo_certidão:
                                    certidão = open(arquivo_certidão, 'rb')
                                    certidão_lida = PyPDF2.PdfReader(
                                        certidão
                                    )
                                    for página_da_certidão in range(
                                            len(certidão_lida.pages)
                                    ):
                                        pagina_lida = certidão_lida.pages[
                                            página_da_certidão
                                        ]
                                        pdf_temporário.add_page(pagina_lida)
                            compilado = open(
                                (
                                    f'{self.comprovantes_de_pagamento}/'
                                    f'Mesclados/{arquivo_pdf[0:-4]}_mesclado.'
                                    f'pdf'
                                ),
                                'wb'
                            )
                            pdf_temporário.write(compilado)
                            compilado.close()
                            pagamento.close()
                            certidão.close()
            messagebox.showinfo(
                DIGITALIZADOS_MESCLADOS[0],
                DIGITALIZADOS_MESCLADOS[1]
            )
            print(nomes_errados)

    def rotina_de_mesclagem(self):
        self.analisa_referencia()
        self.pega_fornecedores()
        self.merge()

    def apaga_imagem(self):
        for emp in self.empresas:
            if not os.path.exists(f'{self.pasta_de_certidões}/{str(emp)}'):
                messagebox.showerror(
                    PASTA_NAO_ENCONTRADA[0],
                    (
                        f'{self.pasta_de_certidões}/{str(emp)}'
                        f'{PASTA_NAO_ENCONTRADA[1]}'
                    )
                )
            os.chdir(f'{self.pasta_de_certidões}/{str(emp)}')
            for arquivo in os.listdir(
                    f'{self.pasta_de_certidões}/{str(emp)}'
            ):
                if arquivo.endswith(".jpg"):
                    os.unlink(
                        f'{self.pasta_de_certidões}/{str(emp)}/{arquivo}'
                    )
