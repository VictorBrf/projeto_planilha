import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    # Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(426, 354, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(348, 453, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(333, 573, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Código Produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(366, 660, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(336, 741, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(348, 825, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão próximo
    pyautogui.click(369, 882, duration=1)
    sleep(3)

    # Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(354, 369, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(342, 465, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Data de validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(339, 555, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(345, 630, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Tamanho
    tamanho = linha[10].value
    pyautogui.click(363, 723, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(348, 744, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(345, 768, duration=1)
    else:
        pyautogui.click(345, 792, duration=1)

    # material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(363, 798, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    # Botão próximo
    pyautogui.click(345, 858, duration=1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(351, 393, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(345, 489, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(336, 582, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(342, 696, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(342, 789, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Botão concluir
    pyautogui.click(339, 846, duration=1)
    # Botão confirmar inclusão
    pyautogui.click(1134, 192, duration=1)
    # iniciar cadastro novamente
    pyautogui.click(918, 627, duration=1)

# Site usado para teste: https://cadastro-produtos-devaprender.netlify.app/index.html
