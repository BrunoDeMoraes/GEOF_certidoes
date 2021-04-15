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
            self.mensagem_de_log_completa(
                f'Execução interrompida!!!\nNão foi possível encontrar o padrão de CNPJ na certidão TST da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.', self.caminho_de_log)
            messagebox.showerror('Esse arquivo não rola!',
                                 f'''Não foi possível encontrar o padrão de CNPJ na certidão TST da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão TST da empresa {emp} inválido.')
        texto = []
        padrao = re.compile('Expedição: (\d\d)/(\d\d)/(\d\d\d\d)')
        emissao_string = padrao.search(certidao)
        try:
            texto.append(emissao_string.group().split()[1])
        except AttributeError:
            self.mensagem_de_log_completa(
                f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de emissão na certidão TST da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.', self.caminho_de_log)
            messagebox.showerror('Esse arquivo não rola!',
                                 f'''Não foi possível encontrar o padrão de data de emissão na certidão TST da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão TST da empresa {emp} inválido.')
        padrao = re.compile('Validade: (\d\d)/(\d\d)/(\d\d\d\d)')
        vencimento_string = padrao.search(certidao)
        try:
            texto.append(vencimento_string.group().split()[1])
        except AttributeError:
            self.mensagem_de_log_completa(
                f'Execução interrompida!!!\nNão foi possível encontrar o padrão de data de vencimento na certidão TST da empresa {emp}.\nO arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.', self.caminho_de_log)
            messagebox.showerror('Esse arquivo não rola!',
                                 f'''Não foi possível encontrar o padrão de data de vencimento na certidão TST da empresa {emp}. O arquivo pode estar corrompido ou ter sofrido atualizações que alteraram sua formatação.''')
            raise Exception(f'Arquivo da certidão TST da empresa {emp} inválido.')
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        self.mensagem_de_log_sem_data(f'   TST   - emissão: {emissao}; válida até: {vencimento}', self.caminho_de_log)
        if validação_de_cnpj in self.lista_de_cnpj:
            self.mensagem_de_log_simples(
                f'   O CNPJ encontrado, {validação_de_cnpj}, pertence à empresa {self.lista_de_cnpj[validação_de_cnpj]}\n', self.caminho_de_log)
        else:
            self.mensagem_de_log_simples(f'   O CNPJ encontrado, {validação_de_cnpj}, não possui correspondência\n', self.caminho_de_log)
        return (data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento), validação_de_cnpj