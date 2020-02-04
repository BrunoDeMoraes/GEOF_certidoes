from certidao2 import Certidao

dia = '31'
mes = '01'
ano = '2020'

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
obj1.pdf_para_jpg_renomear(fornecedores)
obj1.gera_nome(fornecedores)
obj1.apaga_imagem(fornecedores)
