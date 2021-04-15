from certidão import Certidao

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
            self.mensagem_de_log_completa(f'''Execução interrompida!!!
Não foi possível encontrar o padrão de CNPJ na certidão GDF da empresa {emp}.
O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''', self.caminho_de_log)
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
                self.mensagem_de_log_completa(
                    f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de emissão na certidão GDF da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.', self.caminho_de_log)
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
                self.mensagem_de_log_completa(
                    f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de vencimento na certidão GDF da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.', self.caminho_de_log)
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
                self.mensagem_de_log_completa(
                    f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de emissão na certidão GDF da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.', self.caminho_de_log)
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
                self.mensagem_de_log_completa(
                    f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de vencimento na certidão GDF da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.', self.caminho_de_log)
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
        self.mensagem_de_log_sem_data(f'   GDF   - emissão: {emissao}; válida até: {vencimento}', self.caminho_de_log)
        if validação_de_cnpj in self.lista_de_cnpj:
            self.mensagem_de_log_simples(
                f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à empresa {self.lista_de_cnpj[validação_de_cnpj]}\n', self.caminho_de_log)
        else:
            self.mensagem_de_log_simples(f'   O CNPJ encontrado, {validação_de_cnpj}, não possui correspondência\n', self.caminho_de_log)
        return (data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento), validação_de_cnpj