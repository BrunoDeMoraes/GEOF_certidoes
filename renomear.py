from certidao2 import Certidao
import os
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import re
import shutil

dia = '31'
mes = '10'
ano = '2019'
orgaos = ['UNIÃO', 'TST', 'FGTS', 'GDF']
empresasdic = {}


obj1 = Certidao(dia, mes, ano)
lis_ref_cel = obj1.pega_referencia()
if len(lis_ref_cel) == 0:
    raise Exception('Data não encontrada!')
elif len(lis_ref_cel) > 1:
    raise Exception(f'A data especificada foi encontrada nas células {lis_ref_cel} da planilha de pagamentos.'
          f'\nApague os valores duplicados e execute o programa novamente.')
else:
    ref_cel = lis_ref_cel[0]
fornecedores = obj1.pega_fornecedores(ref_cel)
print(fornecedores)

pasta_mae = '//hrg-74977/GEOF/CERTIDÕES/Certidões2'
for emp in fornecedores:
    os.chdir(f'{pasta_mae}/{emp}')
    os.makedirs(f'{pasta_mae}/0 - Renomear/{emp}')
    for arquivo in os.listdir(f'{pasta_mae}/{emp}'):
        if arquivo.endswith('pdf'):
            shutil.copy(f'{pasta_mae}/{emp}/{arquivo}', f'{pasta_mae}/0 - Renomear/{emp}/{arquivo}')

pasta_mae2 = '//hrg-74977/GEOF/CERTIDÕES/Certidões2/0 - Renomear'
for empresa in os.listdir(pasta_mae2):
    origem = f'{pasta_mae2}/{empresa}'
    print(empresa)
    os.chdir(origem)
    for imagem in os.listdir(origem):
        if imagem.endswith(".pdf"):
            pages = convert_from_path(imagem, 300)
            imagem = imagem[:-4]
            for page in pages:
                page.save(f"{imagem}.jpg", "JPEG")
    for imagem in os.listdir(origem):
        if imagem.endswith(".jpg"):
            certidao = pytesseract.image_to_string(Image.open(f'{origem}/{imagem}'), lang = 'por')
            padroes = ['FGTS - CRF', 'JUNTO AO GDF', 'JUSTIÇA DO TRABALHO', 'MINISTÉRIO DA FAZENDA']
            valores = {'FGTS - CRF': 'FGTS', 'JUNTO AO GDF': 'GDF', 'JUSTIÇA DO TRABALHO': 'TST', 'MINISTÉRIO DA FAZENDA': 'UNIÃO'}
            datas = {'FGTS - CRF': 'a (\d\d)/(\d\d)/(\d\d\d\d)', 'JUNTO AO GDF': 'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?'
                                '(Setembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)', 'JUSTIÇA DO TRABALHO': 'Validade: (\d\d)/(\d\d)/(\d\d\d\d)', 'MINISTÉRIO DA FAZENDA': 'Válida até (\d\d)/(\d\d)/(\d\d\d\d)'}
            for frase in padroes:
                if frase in certidao:
                    print(f'certidão {valores[frase]}')
                    data = re.compile(datas[frase])
                    procura = data.search(certidao)
                    datanome = procura.group()
                    print(datanome)
                    separa = datanome.split('/')
                    junta = '-'.join(separa)
                    print(junta)
                    if ':' in junta:
                        retira = junta.split(':')
                        volta = ' '.join(retira)
                        print(volta)
                        junta = volta
                        print(f'tst junta = {junta}')
                    shutil.move(f'{origem}/{imagem[0:-4]}.pdf', f'{valores[frase]} - {junta}.pdf')
    for imagem in os.listdir(origem):
        if imagem.endswith(".jpg"):
            print(f'{imagem}')
            os.unlink(f'{origem}/{imagem}')