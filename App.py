import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informações de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):

    #Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1230,320,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1241,429,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1231,600,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Código do Produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1241,710,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1224,815,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1226,924,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1249,990,duration=1)
    sleep(3)

    #preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1234,349,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1228,461,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1234,573,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1260,680,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Tamanho
    tamanho = linha[10].value
    pyautogui.click(1266,780,duration=1)

    if tamanho == 'Pequeno':
        pyautogui.click(1260,826,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1264,852,duration=1)
    else:
        pyautogui.click(1274,877,duration=1)

    #Material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1226,884,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Botão próximo
    pyautogui.click(1263,966,duration=1)

    #Fabricante
    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1257,381,duration=1)
    pyautogui.hotkey('ctrl','v')

    #País origem
    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1246,486,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Observações
    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1248,591,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Código de barras
    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1247,757,duration=1)
    pyautogui.hotkey('ctrl','v')


    #Localização armazém
    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1249,866,duration=1)
    pyautogui.hotkey('ctrl','v')

    #Botão concluir
    pyautogui.click(1268,938,duration=1)

    #Botão ok
    pyautogui.click(1755,236,duration=1)

    #Adicionar mais um produto
    pyautogui.click(1505,655,duration=1)











# Repetir esses passos para outros campos até preencher campos daquela página
# Clicar em próxima
# Repetir os mesmos passos e ir para a próxima página (página 2)
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir
# Clicar em ok para finalizar o processo
# Clicar no ok mais uma vez na mensagem de confirmação de salvamento no banco de dados
# Clicar em "Adicionar mais um e repetir o processo até finalizar a planilha"

# PyAutoGui (Automação de clicks e teclado)
# Openpyxl  (Leitura e automação de planilhas)