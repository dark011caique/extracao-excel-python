import openpyxl
import pyperclip
import pyautogui
from time import sleep

#https://cadastro-produtos-devaprender.netlify.app/etapa3.html

#https://drive.google.com/drive/folders/1O5981-PvouJkue2s09HUuagr7A-pWJly


worknook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = worknook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1106, 279, duration=1) 
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1033,405,duration=1 )
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1035,557,duration=1 )
    pyautogui.hotkey('ctrl','v')

    codigo = linha[3].value
    pyperclip.copy(codigo)
    pyautogui.click(1131,656,duration=1 )
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1096,764,duration=1 )
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1094,878,duration=1 )
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1059,942,duration=1)
    sleep(5)


    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1076,308,duration=1 )
    pyautogui.hotkey('ctrl','v')

    quantidade = linha[7].value
    pyperclip.copy(quantidade)
    pyautogui.click(1077,421,duration=1 )
    pyautogui.hotkey('ctrl','v')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(1078,520,duration=1 )
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1053,636,duration=1 )
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyautogui.click(1169,739,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(1125, 779,duration=1 )
    elif tamanho =='Médio':
        pyautogui.click(1124,819,duration=1 )
    else:
        pyautogui.click(1125,852,duration=1 )

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1055,845,duration=1 )
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1059,942,duration=1)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1088,333,duration=1 )
    pyautogui.hotkey('ctrl','v')

    pais = linha[13].value
    pyperclip.copy(pais)
    pyautogui.click(1096,444,duration=1 )
    pyautogui.hotkey('ctrl','v')

    oberservacao = linha[14].value
    pyperclip.copy(oberservacao)
    pyautogui.click(1084,562,duration=1 )
    pyautogui.hotkey('ctrl','v')

    codigo_barra = linha[15].value
    pyperclip.copy(codigo_barra)
    pyautogui.click(1082,716,duration=1 )
    pyautogui.hotkey('ctrl','v')

    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.click(1079,822,duration=1 )
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1047,899,duration=1)
    sleep(3)

    pyautogui.click(1661,209,duration=1)
    sleep(3)

    pyautogui.click(1376,609,duration=1)
    sleep(3)