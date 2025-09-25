import pyautogui as pg
from time import sleep 
import pyperclip as pc
import pandas as pd
from pyautogui import ImageNotFoundException
import os

CAMINHO_IMAGENS = '1- Trancar Conta Caixa'
data = '19/09/2025'

df = pd.read_excel('1- Trancar Conta Caixa.xlsx', dtype={'Codcxa': str,'Coligada': str})

valores_vistos = set()

ocorrencias_coligada = df['Coligada'].value_counts()

pg.PAUSE = 1

# Função para buscar uma imagem até encontrar
def encontrar_imagem(nome_imagem, confiança=0.7, grayscale=True):
    caminho_completo = os.path.join(CAMINHO_IMAGENS, nome_imagem)
    while True:
        try:
            # Tenta encontrar a imagem na tela
            resultado = pg.locateOnScreen(caminho_completo, grayscale=grayscale, confidence=confiança)
        
            # Se a imagem for encontrada, retorna as coordenadas
            if resultado:
                x, y, largura, altura = resultado
                return x, y, largura, altura
        except ImageNotFoundException:
            # Caso não encontre, espera 1 segundo e tenta novamente
            print(f"Imagem {nome_imagem} não encontrada. Tentando novamente...")
            sleep(1)

# Função para fazer a automação considerando que é a primeira vez entrando na coligada
def primeira_vez():
    x, y, largura, altura = encontrar_imagem('1-Contexto.png', 0.7)
    pg.click(x + largura / 2, y + altura / 2)

    x, y, largura, altura = encontrar_imagem('2-Coligada.png', 0.7)
    pg.click(x + largura / 2, y + altura / 2)

    x, y, largura, altura = encontrar_imagem('3-Avancar.png', 0.7)
    pg.click(x + largura / 2, y + altura / 2)
     
    x, y, largura, altura = encontrar_imagem('4-Inserir coligada.png', 0.7)
    pg.click(x + largura / 2, y + altura / 1.4)

    pg.write(coligada)
    pg.press('tab')
    pg.click(1090,720)
    pg.click(1090,720)
    sleep(8)

    x, y, largura, altura = encontrar_imagem('5-Movimentacoes Bancarias.png', 0.7, grayscale=False)
    pg.click(x + largura / 2, y + altura / 1.4)

    x, y, largura, altura = encontrar_imagem('6-Contas-Caixa.png', 0.7)
    pg.click(x + largura / 2, y + altura / 2)

    x, y, largura, altura = encontrar_imagem('7-ContaCaixaRobo.png', 0.7)
    pg.doubleClick(x + largura / 2, y + altura / 2)

    pc.copy(conta_caixa) # copiar conta caixa 
    pg.hotkey('ctrl','v') # colar a conta caixa
    pg.press('enter') 
    pg.doubleClick(195, 388) # para abrir o lançamento no Totvs

    x, y, largura, altura = encontrar_imagem('8-Dados Adicionais.png', 0.7)
    pg.click(x + largura / 2, y + altura / 2) # ir para dados adicionais

    pg.doubleClick(667,501) # para alterar a data
    pc.copy(data)
    pg.hotkey('ctrl','v')
    pg.click(1070,830) # para fechar a janela
    sleep(1)
    valores_vistos.add(coligada)

# Função para dar continuidade ao processo de automação sem precisar repetir os primeiros passos que ocorrem na primeira vez
def nem_nem():
    pc.copy(conta_caixa) # copiar conta caixa

    x, y, largura, altura = encontrar_imagem('9-Filtro ContaCaixa.png', 0.7)
    pg.click(x + largura / 2, y + altura / 2) # ir para dados adicionais # clicar na janela de conta caixa no totvs

    pg.hotkey('ctrl','v') # colar a conta caixa 
    pg.click(993,645) # clicar no ok para seguir
    pg.doubleClick(195, 388) # para abrir o lançamento no Totvs

    x, y, largura, altura = encontrar_imagem('8-Dados Adicionais.png', 0.7)
    pg.click(x + largura / 2, y + altura / 2)

    pg.doubleClick(667,501) # para alterar a data
    pc.copy(data)
    pg.hotkey('ctrl','v')
    pg.click(1070,830) # para fechar a janela


pg.alert("O código irá começar")


for i, coligada in enumerate(df['Coligada']):
    conta_caixa = df.loc[i,"Codcxa"]
    
    # Verificando se é a primeira e última vez entrando na coligada
    if coligada not in valores_vistos and ocorrencias_coligada[coligada] == 1: 
        print(f'Primeira e última vez entrando na coligada: {coligada}, conta caixa: {conta_caixa}')
        primeira_vez()
        pg.hotkey('ctrl','w')
         
    elif coligada not in valores_vistos:
        print(f'Primeira vez entrando na coligada: {coligada}, conta caixa: {conta_caixa}')        
        primeira_vez()

    # Verificando se é a última vez entrando na coligada
    elif ocorrencias_coligada[coligada] == 1: 
        print(f'Última vez entrando na coligada {coligada}, conta caixa: {conta_caixa}')
        nem_nem()
        sleep(1)
        pg.hotkey('ctrl','w')

    else:   
        print(f'Nem primeira nem ultima vez entrando na coligada: {coligada}, conta caixa: {conta_caixa}')
        nem_nem()
        
    ocorrencias_coligada[coligada] -= 1


pg.alert("O código foi finalizado")