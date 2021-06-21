#strings simples em ordem alfabética

ANALISADOS = '\nFornecedores analisados:'

CONFERENCIA = '\nRESULTADO DA CONFERÊNCIA:'

CHECA_URL_0 = (
    'O arquivo xlsx selecionado como fonte foi apagado, removido ou não exist'
    'e.\n\nClique em Configurações>>Caminhos>>Fonte de dados XLSX e selecione'
    ' um arquivo xlsx que atenda aos critéri os necessários para o processame'
    'nto.'
)

CHECA_URL_1 = (
    'A pasta apontada como fonte para certidões foi apagada, removida ou não '
    'existe.\n\nClique em Configurações>>Caminhos>>Pasta de certidões e selec'
    'ione uma pasta que contenha as certidões que devem ser analisadas.'
)

CHECA_URL_2 = (
    'A pasta apontada como fonte e destino para logs foi apagada, removida ou'
    ' não existe.\n\nClique em Configurações>>Caminhos>>Pasta de logs e selec'
    'ione uma pasta onde os logs serão criados.'
)

CHECA_URL_3 = (
    'A pasta apontada como fonte de comprovantes de pagamento foi apagada, re'
    'movida ou não existe.\n\nClique em Configurações>>Caminhos>>Comprovantes'
    ' de pagamento e selecione uma pasta que contenha os comprovantes de paga'
    'mento.'
)

CHECA_URL_4 = (
    'A pasta apontada como destino de certidões para pagamento foi apagada, r'
    'emovida ou não existe.\n\nClique em Configurações>>Caminhos>>Certidões p'
    'ara pagamento e selecione uma pasta para direcionar as certidões do paga'
    'mento.'
)

CNPJ_EMPRESA = '- o CNPJ encontrado pertence à empresa'

CNPJ_NULO = '- o CNPJ encontrado não possui correspondência'

DATA_NACIONAL = '%d/%m/%Y'

INICIO_DA_EXECUCAO = (
    '\n======================================================================'
    '========================================================================'
    '\n\nInício da execução'
)

INICIO_DA_ANALISE = (
    '\nInicio da conferência de datas de emissão e vencimento:'
)

LINHA_FINAL = (
    '\n\n===================================================================='
    '========================================================================'
    '=\n'
)

PADRAO_CNPJ = '(\d\d).(\d\d\d).(\d\d\d)/(\d\d\d\d)-(\d\d)'

PADRAO_FGTS = '(\d\d)/(\d\d)/(\d\d\d\d) a (\d\d)/(\d\d)/(\d\d\d\d)'

PADRAO_GDF = (
    'Brasília, (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)'
    '?(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?('
    'fevereiro)?(março)?(abril)?(maio)?(junho)?(julho)?(agosto)?(setembro)?(o'
    'utubro)?(novembro)?(dezembro)? de (\d\d\d\d)'
)

PADRAO_GDF_2 = (
    'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho'
    ')?(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(janeiro)?'
    '(fevereiro)?(março)?(abril)?(maio)?(junho)?(julho)?(agosto)?(setembro)?('
    'outubro)?(novembro)?(dezembro)? de (\d\d\d\d)'
)

PADRAO_GDF_3 = 'Certidão emitida via internet em (\d\d)/(\d\d)/(\d\d\d\d)'

PADRAO_GDF_4 = (
    'Válida até (\d)?(\d\d)? de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?'
    '(Junho)?(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(jan'
    'eiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?(julho)?(agosto)?(setem'
    'bro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)'
)

PASTA_CRIADA = (
    '\nPasta para inclusão de arquivos de pagamento criada com sucesso.\n',
    'Número de novas pastas criadas: '
)

PASTA_DE_PAGAMENTO = (
    'FICA CALMO!!!',
    ('Já existe pasta contendo certidões para pagamento na data informada.\nS'
     'e deseja fazer nova transferência apague o diretório: ')
)

PASTA_LOCALIZADA = (
    '\nPasta para inclusão de arquivos de pagamento localizada.\n'
)

PENDENCIAS = '\n\nCERTIDÕES QUE DEVEM SER ATUALIZADAS:\n'

REFERENCIA = 'Referência encontrada na célula '


TEXTO_ANALISAR = (
    'Utilize esta opção para identificar quais certidões devem ser atualizada'
    's ou se há requisitos a cumprir para a devida execução da análise.'
)

TEXTO_CRIA_ESTRUTURA = (
    'Se deseja criar toda a estrutura de pastas de trabalho necessárias para '
    'o\ncorreto funcionamento do programa na pasta que contém o arquivo princ'
    'ipal,\nclique em "Criar estrutura". Em seguida, selecione manualmente ca'
    'da caminho.'
)

TEXTO_MESCLA_ARQUIVOS = (
    'Após o pagamento utilize esta opção para mesclar os comprovantes de paga'
    'mento digitalizados com suas respectivas certidões.'
)

TEXTO_PRINCIPAL = (
    '    Indique a data limite pretendida para o próximo pagamento e em segui'
    'da escolha uma das seguintes opções:    '
)

TEXTO_RENOMEAR = (
    'Após atualizar as certidões, selecione uma das opções para padronizar os'
    ' nomes dos\narquivos e em seguida faça nova análise para certificar que '
    'está tudo OK.'
)

TEXTO_TRANSFERE_ARQUIVOS = (
    'Esta opção transfere as certidões que validam o pagamento para uma pasta'
    ' identificada pela data.\nEsse passo deve ser executado logo após a anál'
    'ise definitiva antes do pagamento.'
)

TITULO_DA_INTERFACE = 'GEOF - Analisador de certidões'

#tuplas de strings por ordem alfabética

ANALISE_EXECUTADA = (
    'Analisou, miserávi!',
    'Processo de análise de certidões executado com sucesso!'
)

ARQUIVO_CORROMPIDO = (
    'Arquivo zuado!!!',
    f"Arquivo corrompido: ")


ARQUIVO_INEXISTENTE = (
    'O arquivo selecionado não existe.',
    'Esse arquivo é invenção da sua cabeça, parça!',
    'O arquivo selecionado não existe ou já foi renomeado!'
)

ARQUIVOS_NAO_SELECIONADOS = (
    'Se não selecionar os arquivos, não vai rolar!',
    'Selecione os arquivos que deseja renomear'
)

ARQUIVOS_INVALIDOS = (
    'arquivo da certidão FGTS da empresa inválido',
    'arquivo da certidão GDF da empresa inválido',
    'arquivo da certidão TST da empresa inválido',
    'arquivo da certidão UNIÃO da empresa inválido'
)

ATUALIZAR_CAMINHOS = (
    'Vc sabe o que está fazendo?',
    'Deseja alterar a configuração dos caminhos de pastas e arquivos?'
)

ATUALIZAR_XLSX = (
    'Problema com o xlsx',
    (
        'O arquivo fonte de dados XLSX parece não ter sido atualizado correta'
        'mente.\n\nTente atualizar a planilha FORNECEDORES usando a oção de c'
        'olagem  que insere apenas "Valores"'
    )
)

CAMINHOS_ATUALIZADOS = (
    'Fez porque quis!',
    'Caminhos de pastas e arquivos utilizados pelo sistema atualizados.'
)

CERTIDOES_FALTANDO = (
    ', não foram encontradas as certidões:',
    f'Adicione as certidões às respectivas pastas informadas e execute novame'
    f'nte o programa.',
    f'Tá faltando coisa, mano!',
    f'Algumas certidões não foram encontradas!\nConsulte o arquivo de log, re'
    f'solva as pendências indicadas e então execute novamente a análise.'
)


CERTIDOES_TRANSFERIDAS = (
    'Transferiu, miserávi!',
    'As certidões que validam o pagamento foram transferidas com sucesso!'
)

CNPJ_NAO_ENCONTRADO = (
    'Esse arquivo não rola!',
    (
        'Não foi possível encontrar o padrão de CNPJ na certidão.\nO arquivo '
        'pode estar corrompido ou ter sofrido atualizações que alteraram sua '
        'formatação.'
    )
)

CRIANDO_IMAGENS = (
    '========================================================================'
    '========================================================================'
    '\nCriando imagens:\n',
    '\nImagens criadas com sucesso!\n'
)

DADOS_DO_FORNECEDOR_COM_ERRO = (
    'Dados do fornecedor estão zuados!',
    f'Não foi possível localizar o CNPJ da empresa na planilha FORNECEDORES d'
    f'o arquivo xlsx utilizado como fonte de dados.\n\nVerifique se há regist'
    f'ro de CNPJ para a empresa ou se o nome informado na planilha PAGAMENTO '
    f'é idêntico ao inserido na planilha FORNECEDORES.'
)

DATAS_MULTIPLAS = (
    'Me ajuda a te ajudar!',
    (
        'Data informada em multiplicidade.\nA data especificada como limite p'
        'ara pagamento foi encontrada nas células:'
    ),
    (
        '\nApague as células com valores duplicados na planilha de pagamentos'
        ' e execute o programa novamente.'
    )
)

DATA_NAO_ENCONTRADA = (
    'Me ajuda a te ajudar!',
    (
        '\nA data informada não foi encontrada na lista de datas para pagamen'
        'to ou não existe!'
    )
)

DIGITALIZADOS_MESCLADOS = (
    'Mesclou, miserávi!!!',
    'Digitalizações de pagamentos e respectivas certidões mescladas com sucesso!'
)

EMISSAO_VENCIMENTO = (
    'Esse arquivo não rola!',
    ('Não foi possível encontrar o padrão de data de emissão e vencimento na '
     'certidão  da empresa .\nO arquivo pode estar corrompido ou ter sofrido '
     'atualizações que alteraram sua formatação.'
     )
)

IDENTIFICADOR_DE_CERTIDAO = (
    'FGTS - CRF',
    'Brasília,',
    'JUSTIÇA DO TRABALHO',
    'MINISTÉRIO DA FAZENDA',
    'GOVERNO DO DISTRITO FEDERAL'
)

LOG_INEXISTENTE = (
    'Me ajuda a te ajudar!',
    'Não existe log para a data informada.'
)

OPCOES_DE_RENOMEACAO = (
    'Selecione uma opção',
    'Renomear arquivos',
    'Renomear todos os arquivos de uma pasta',
    'Renomear todas as certidões da lista de pagamento'
)

OPCOES_DE_RENOMEAR_VAZIA = (
    'Tem que escolher, fi!',
    'Nenhuma opção selecionada!'
)

ORGAOS = ('UNIÃO', 'TST', 'FGTS', 'GDF')

PASTA_DE_MESCLAGEM_EXISTENTE = (
    'FICA CALMO!!!',
    f'Já existe pasta para mesclagem na data informada!\nSe deseja fazer nova'
    f' mesclagem apague o diretório: '
)


PASTA_NAO_ENCONTRADA = (
    'Tem essa pasta aí não, locão!',
    (
        ' ainda não existe.\n\nAntes de tentar renomear as certidões, execute'
        ' a opção "Analisar certidões". A referida opção criará as pastas nec'
        'essárias e indicará o que precisa ser atualizado antes do processo d'
        'e renomeação.'
    )
)

PASTA_NAO_SELECIONADA = (
    'Se não selecionar a pasta, não vai rolar!',
    'Selecione uma pasta para renomear todos os seu arquivos de certidões.'
)

PLANILHAS = (
    'FORNECEDORES',
    'PAGAMENTO',
    'Esse arquivo não rola!',
    'O arquivo xlsx selecionado como fonte não possui as planilhas necessária'
    's para o processamento solicitado.\n\nClique em Configurações>>Caminhos>'
    '>Fonte de dados XLSX e selecione um arquivo xlsx que atenda aos critério'
    's necessários para o processamento.'
)

RENOMEACAO_EXECUTADA = (
    'Renomeou, miserávi!',
    ('Todas as certidões da listagem de pagamento foram renomeadas com sucess'
     'o!'),
    '\nProcesso de renomeação de certidão executado com sucesso!',
    'Todas as certidões da pasta selecionada foram renomeadas co'
    'm sucesso!'
)

#dicionários de strings

IDENTIFICADOR_TRADUZIDO = {
    'FGTS - CRF': 'FGTS',
    'Brasília,': 'GDF',
    'JUSTIÇA DO TRABALHO': 'TST',
    'MINISTÉRIO DA FAZENDA': 'UNIÃO',
    'GOVERNO DO DISTRITO FEDERAL': 'GDF'
}

IDENTIFICADOR_DE_VALIDADE = {
    'FGTS - CRF': 'a (\d\d)/(\d\d)/(\d\d\d\d)',

    'Brasília,': (
        'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(J'
        'unho)?(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)? d'
        'e (\d\d\d\d)'),

    'JUSTIÇA DO TRABALHO': 'Validade: (\d\d)/(\d\d)/(\d\d\d\d)',

    'MINISTÉRIO DA FAZENDA': 'Válida até (\d\d)/(\d\d)/(\d\d\d\d)',

    'GOVERNO DO DISTRITO FEDERAL': (
        'Válida até (\d\d) de (Janeiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(J'
        'unho)?(Julho)?(Agosto)?(Setembro)?(Outubro)?(Novembro)?(Dezembro)?(j'
        'aneiro)?(fevereiro)?(março)?(abril)?(maio)?(junho)?(julho)?(agosto)?'
        '(setembro)?(outubro)?(novembro)?(dezembro)? de (\d\d\d\d)')
}

IDENTIFICADOR_DE_VALIDADE_2 = {
    'GOVERNO DO DISTRITO FEDERAL': (
        'Válida até (\d) de (janeiro)?(fevereiro)?(março)?(abril)?(maio)?(jun'
        'ho)?(julho)?(agosto)?(setembro)?(outubro)?(novembro)?(dezembro)?(Jan'
        'eiro)?(Fevereiro)?(Março)?(Abril)?(Maio)?(Junho)?(Julho)?(Agosto)?(S'
        'etembro)?(Outubro)?(Novembro)?(Dezembro)? de (\d\d\d\d)')
}

MESES = {
    'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04',
    'Maio': '05', 'Junho': '06', 'Julho': '07', 'Agosto': '08',
    'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'
}

MESES2 = {
    'janeiro': '01', 'fevereiro': '02', 'março': '03', 'abril': '04',
    'maio': '05', 'junho': '06', 'julho': '07', 'agosto': '08',
    'setembro': '09', 'outubro': '10', 'novembro': '11', 'dezembro': '12',
    'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04',
    'Maio': '05', 'Junho': '06', 'Julho': '07', 'Agosto': '08',
    'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'
}