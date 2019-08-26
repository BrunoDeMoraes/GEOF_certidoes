from pdf2image import convert_from_path
from PIL import Image
import openpyxl
import os
import pytesseract
import re
import time

class Certidao:
    def __init__(self, dia, mes, ano):
        self.dia = dia
        self.mes = mes
        self.ano = ano
        self.wb = openpyxl.load_workbook(r'D:\Leiturapdf\Matrix PDPAS 2019 - HRG.xlsx')
        self.pag = self.wb['PAGAMENTO']
        self.listareferencia = []
        self.referencia = 0
        self.datapag = 'CERTIDÕES PARA {}/{}/{}'.format(self.dia, self.mes, self.ano)
        self.empresas = []
        self.pdf_dir = r'\\hrg-74977\GEOF\CERTIDÕES\Certidões - Bruno_teste\\'

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
            self.empresas.append(empresa[0])
            desloca += 1
        return self.empresas

    def cria_diretorio(self, fornecedores):
        novos_dir = []
        for emp in fornecedores:
            if os.path.isdir(self.pdf_dir + str(emp)):
                continue
            else:
                os.makedirs(self.pdf_dir + str(emp) + '\Vencidas')
                os.makedirs(self.pdf_dir + str(emp) + '\Imagens')
                novos_dir.append(emp)
        print(f'Foram criadas {len(novos_dir)} novas pastas {novos_dir}.')

    def certidoes_n_encontradas(self, fornecedores, orgaos):
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

    def pdf_para_jpg(self, fornecedores, orgaos):
        for emp in fornecedores:
            os.chdir(self.pdf_dir + str(emp))
            for pdf_file in os.listdir(self.pdf_dir + str(emp)):
                if pdf_file.endswith(".pdf") and pdf_file.split()[0] in orgaos:
                    pages = convert_from_path(pdf_file, 300)
                    pdf_file = pdf_file[:-4]
                    for page in pages:
                        page.save("{}.jpg".format(pdf_file), "JPEG")


class Uniao(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(self.pdf_dir + str(emp))
        for imagem in os.listdir(self.pdf_dir + str(emp)):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'UNIAO':
                certidao = pytesseract.image_to_string(
                    Image.open(r'\\hrg-74977\GEOF\CERTIDÕES\Certidões - Bruno_teste\{}\{}'.format(emp, imagem)),
                    lang='por')
                return certidao

    def confere_data(self, certidao):
        texto = []
        padrao = re.compile('do dia (\d\d)/(\d\d)/(\d\d\d\d)')
        emissao_string = padrao.search(certidao)
        print(emissao_string.group())
        texto.append(emissao_string.group().split()[2])
        padrao = re.compile('Válida até (\d\d)/(\d\d)/(\d\d\d\d)')
        vencimento_string = padrao.search(certidao)
        print(vencimento_string.group())
        texto.append(vencimento_string.group().split()[2])
        print(texto)
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        if data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento:
            return True
        else:
            return False

class Tst(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(self.pdf_dir + str(emp))
        for imagem in os.listdir(self.pdf_dir + str(emp)):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'TST':
                certidao = pytesseract.image_to_string(
                    Image.open(r'\\hrg-74977\GEOF\CERTIDÕES\Certidões - Bruno_teste\{}\{}'.format(emp, imagem)),
                    lang='por')
                return certidao

    def confere_data(self, certidao):
        texto = []
        padrao = re.compile('Expedição: (\d\d)/(\d\d)/(\d\d\d\d)')
        emissao_string = padrao.search(certidao)
        print(emissao_string.group())
        texto.append(emissao_string.group().split()[1])
        padrao = re.compile('Validade: (\d\d)/(\d\d)/(\d\d\d\d)')
        vencimento_string = padrao.search(certidao)
        print(vencimento_string.group())
        texto.append(vencimento_string.group().split()[1])
        print(texto)
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        if data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento:
            return True
        else:
            return False

class Fgts(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(self.pdf_dir + str(emp))
        for imagem in os.listdir(self.pdf_dir + str(emp)):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'FGTS':
                certidao = pytesseract.image_to_string(
                    Image.open(r'\\hrg-74977\GEOF\CERTIDÕES\Certidões - Bruno_teste\{}\{}'.format(emp, imagem)),
                    lang='por')
                return certidao

    def confere_data(self, certidao):
        texto = []
        padrao = re.compile('(\d\d)/(\d\d)/(\d\d\d\d) a (\d\d)/(\d\d)/(\d\d\d\d)')
        emissao_string = padrao.search(certidao)
        print(emissao_string.group())
        texto.append(emissao_string.group().split()[0])
        vencimento_string = padrao.search(certidao)
        texto.append(vencimento_string.group().split()[2])
        print(texto)
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        if data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento:
            return True
        else:
            return False

class Gdf(Certidao):
    def __init__(self, dia, mes, ano):
        super().__init__(dia, mes, ano)

    def pega_string(self, emp):
        os.chdir(self.pdf_dir + str(emp))
        for imagem in os.listdir(self.pdf_dir + str(emp)):
            if imagem.endswith(".jpg") and imagem.split()[0] == 'GDF':
                certidao = pytesseract.image_to_string(
                    Image.open(r'\\hrg-74977\GEOF\CERTIDÕES\Certidões - Bruno_teste\{}\{}'.format(emp, imagem)),
                    lang='por')
                return certidao

    def confere_data(self, certidao):
        texto = []
        meses = {'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04', 'Maio': '05', 'Junho': '06',
                 'Julho': '07', 'Agosto': '08', 'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'}
        padrao = re.compile('Brasília, (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                            '(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)')
        emissao_string = padrao.search(certidao)
        print(emissao_string.group())
        datasplit = [emissao_string.group().split()[1], meses[emissao_string.group().split()[3]],
                     emissao_string.group().split()[5]]
        texto.append('/'.join(datasplit))
        padrao = re.compile('Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                            '(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)')
        vencimento_string = padrao.search(certidao)
        print(vencimento_string.group())
        datasplit2 = [vencimento_string.group().split()[2], meses[vencimento_string.group().split()[4]],
                     vencimento_string.group().split()[6]]
        texto.append('/'.join(datasplit2))
        print(texto)
        emissao = texto[0]
        vencimento = texto[1]
        data_de_emissao = time.strptime(emissao, "%d/%m/%Y")
        data_de_vencimento = time.strptime(vencimento, "%d/%m/%Y")
        payday = f'{self.dia}/{self.mes}/{self.ano}'
        data_do_pagamento = time.strptime(payday, "%d/%m/%Y")
        if data_do_pagamento >= data_de_emissao and data_do_pagamento <= data_de_vencimento:
            return True
        else:
            return False

