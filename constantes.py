#strings em ordem alfabética

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

INICIO_DA_EXECUCAO = (
    '\n======================================================================'
    '========================================================================'
    '========================================================================'
    '\n\nInício da execução'
)

LINHA_FINAL = (
    '\n\n===================================================================='
    '========================================================================'
    '========================================================================'
    '======================================================================\n'
)

PENDENCIAS = '\n\nCERTIDÕES QUE DEVEM SER ATUALIZADAS:\n'

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


#listas de strings por ordem alfabética

ANALISE_EXECUTADA = [
    'Analisou, miserávi!',
    'Processo de análise de certidões executado com sucesso!'
]

ATUALIZAR_CAMINHOS = [
'Vc sabe o que está fazendo?',
'Deseja alterar a configuração dos caminhos de pastas e arquivos?'
]

CAMINHOS_ATUALIZADOS = [
'Fez porque quis!',
'Caminhos de pastas e arquivos utilizados pelo sistema atualizados.'
]

LOG_INEXISTENTE = [
    'Me ajuda a te ajudar!',
    'Não existe log para a data informada.']

OPCOES_DE_RENOMEACAO = [
    'Selecione uma opção',
    'Renomear arquivos',
    'Renomear todos os arquivos de uma pasta',
    'Renomear todas as certidões da lista de pagamento',
    'Tem que escolher, fi!',
    'Nenhuma opção selecionada!'
]

RENOMEACAO_EXECUTADA = [
    'Renomeou, miserávi!',
    ('Todas as certidões da listagem de pagamento foram renomeadas com sucess'
     'o!')
]
