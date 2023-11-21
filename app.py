import openpyxl
import pyperclip
import pyautogui
from time import sleep
#Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

#CADASTRAR PRIMEIRA PAGINA
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(799,228,duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(796,302,duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(790,425,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo = linha[3].value
    pyperclip.copy(codigo)
    pyautogui.click(791,505,duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(794,578,duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(791,658,duration=1)
    pyautogui.hotkey('ctrl','v')

    #click no botao proximo
    pyautogui.click(795,704,duration=1)

    sleep(5)

    #segunda pagina 
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(796,248,duration=1)
    pyautogui.hotkey('ctrl','v')

    qtd = linha[7].value
    pyperclip.copy(qtd)
    pyautogui.click(791,325,duration=1)
    pyautogui.hotkey('ctrl','v')

    data = linha[8].value
    pyperclip.copy(data)
    pyautogui.click(793,404,duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(795,474,duration=1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyperclip.copy(tamanho)
    if tamanho == 'Pequeno':
        pyautogui.click(800,559,duration=1)
        pyautogui.click(801,583,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(800,559,duration=1)
        pyautogui.click(799,607,duration=1)
    elif tamanho == "Grande":
        pyautogui.click(800,559,duration=1)
        pyautogui.click(798,628,duration=1)

    material = linha[11].value
    pyperclip.copy(material)        
    pyautogui.click(792,636,duration=1)
    pyautogui.hotkey('ctrl','v')

    #click no botao proximo
    pyautogui.click(803,688,duration=1)

    sleep(5)

    #terceira pagina
    fabricante = linha[12].value
    pyperclip.copy(fabricante)  
    pyautogui.click(797,263,duration=1)
    pyautogui.hotkey('ctrl','v')

    pais = linha[13].value
    pyperclip.copy(pais)
    pyautogui.click(799,345,duration=1)
    pyautogui.hotkey('ctrl','v')

    observacao = linha[14].value
    pyperclip.copy(observacao)
    pyautogui.click(798,415,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo = linha[15].value
    pyperclip.copy(codigo)
    pyautogui.click(796,540,duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao = linha[16].value
    pyperclip.copy(localizacao)
    pyautogui.click(797,618,duration=1)
    pyautogui.hotkey('ctrl','v')
    
    #click no botao concluir
    pyautogui.click(805,672,duration=1)
 
    #click pop up
    sleep(2)
    pyautogui.click(1198,178,duration=1)

    #click em adicionar mais um
    sleep(2)
    pyautogui.click(1029,467,duration=1)


#site para teste da aplicação:
#https://cadastro-produtos-devaprender.netlify.app/
