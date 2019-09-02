import os
import shutil

class Diretorio:
    def __init__(self, pasta_mae):
        self.pasta_mae = pasta_mae
        self.lista_de_arquivos = []

    def lista_pastas(self):
        lista_de_pastas = []
        for emp in os.listdir(self.pasta_mae):
            lista_de_pastas.append(f'{self.pasta_mae}/{emp}')
        return lista_de_pastas

    def lista_arquivos(self, lista_de_pastas):
        for caminho in lista_de_pastas:
            if os.path.isdir(caminho):
                for arquivo in os.listdir(caminho):
                    self.lista_de_arquivos.append(f'{caminho}/{arquivo}')
        return self.lista_de_arquivos

    def renomeia_arquivos(self, lista_de_arquivos, data):
        orgaos = ['UNIAO', 'FGTS', 'TST', 'GDF', 'UNIÃO']
        for arquivo in lista_de_arquivos:
            separa = arquivo.split('/')[0:-1]
            pasta = '/'.join(separa) + '/'
            for orgao in orgaos:
                if not arquivo.endswith('pdf') or 'MERGE' in arquivo.upper():
                    continue
                elif orgao in arquivo.upper():
                    if orgao == 'UNIÃO':
                        orgao = 'UNIAO'
                    shutil.move(f'{arquivo}', f'{pasta}{orgao} - {separa[-1]} - {data}.pdf')