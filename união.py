import os
import time
from tkinter import *
from tkinter import messagebox

import pytesseract
from PIL import Image

from certidão import Certidao
from constantes import ARQUIVOS_INVALIDOS
from constantes import CNPJ_EMPRESA
from constantes import CNPJ_NULO
from constantes import DATA_NACIONAL
from constantes import EMISSAO_VENCIMENTO
from constantes import PADRAO_CNPJ
from constantes import PADRAO_UNIAO
from constantes import CNPJ_NAO_ENCONTRADO


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
        padrão_cnpj = re.compile(PADRAO_CNPJ)
        try:
            validação_de_cnpj = padrão_cnpj.search(certidao).group()
        except AttributeError:
            self.mensagem_de_log_completa(
                f'UNIÃO - {emp}\n{CNPJ_NAO_ENCONTRADO[1]}',
                self.caminho_de_log
            )
            messagebox.showerror(
                CNPJ_NAO_ENCONTRADO[0],
                f'UNIÃO - {emp}\n{CNPJ_NAO_ENCONTRADO[1]}'
            )
            raise Exception(f'{emp} - {ARQUIVOS_INVALIDOS[1]}.')
        texto = []
        padrao = re.compile(PADRAO_UNIAO[0])
        emissao_string = padrao.search(certidao)
        try:
            texto.append(emissao_string.group().split()[2])
        except AttributeError:
            self.mensagem_de_log_completa(
                f'{emp} - UNIÃO\n{EMISSAO_VENCIMENTO[1]}',
                self.caminho_de_log
            )
            messagebox.showerror(
                EMISSAO_VENCIMENTO[0],
                f'{emp} - UNIÃO\n{EMISSAO_VENCIMENTO[1]}'
            )
            raise Exception(f'{emp} - UNIÃO\n{EMISSAO_VENCIMENTO[1]}')

        padrao = re.compile(PADRAO_UNIAO[1])
        vencimento_string = padrao.search(certidao)
        try:
            texto.append(vencimento_string.group().split()[2])
        except AttributeError:
            self.mensagem_de_log_completa(
                f'{emp} - UNIÃO\n{EMISSAO_VENCIMENTO[1]}',
                self.caminho_de_log
            )
            messagebox.showerror(
                EMISSAO_VENCIMENTO[0],
                f'{emp} - UNIÃO\n{EMISSAO_VENCIMENTO[1]}'
            )
            raise Exception(f'{emp} - UNIÃO\n{EMISSAO_VENCIMENTO[1]}')
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, DATA_NACIONAL)
        data_de_vencimento = time.strptime(vencimento, DATA_NACIONAL)
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, DATA_NACIONAL)
        self.mensagem_de_log_sem_data(
            f'    UNIÃO - emissão: {emissao}; válida até: {vencimento}',
            self.caminho_de_log
        )
        if validação_de_cnpj in self.lista_de_cnpj:
            self.mensagem_de_log_simples(
                (
                    f'    {validação_de_cnpj} {CNPJ_EMPRESA} '
                    f'{self.lista_de_cnpj[validação_de_cnpj]}\n'),
                self.caminho_de_log
            )
        else:
            self.mensagem_de_log_simples(
                (
                    f'    {validação_de_cnpj} {CNPJ_NULO}\n'
                ),
                self.caminho_de_log
            )
        return (
            (data_de_emissao <= data_do_pagamento <= data_de_vencimento),
            validação_de_cnpj
        )
