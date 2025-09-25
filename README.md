# Automacao_Trancar_Conta_Caixa
Automação de Trancamento de Contas Caixa (TOTVS)
Este script foi desenvolvido para automatizar uma tarefa repetitiva e diária no sistema TOTVS: o trancamento de contas caixa para diferentes coligadas. A automação reduz o esforço manual e o risco de erros durante o processo.
_____________________________________________________________________________
O que o código faz:
1- *Lê uma planilha Excel* com as coligadas e suas respectivas contas caixa.
2-*Identifica*, via imagem na tela (com `pyautogui`), os elementos do sistema TOTVS** necessários para navegar até o lançamento de contas caixa.
3- *Altera a data do lançamento* para o valor definido no início do script.
4- *Gerencia o fluxo para cada coligada*, considerando:
   - Se é a *primeira vez* acessando a coligada
   - Se é a *última vez*
   - Ou se é apenas uma visita intermediária
_____________________________________________________________________________
Lógica aplicada:
- *`valores_vistos = set()`* -> Controla quais coligadas já foram acessadas.

- *`ocorrencias_coligada = df['Coligada'].value_counts()`*  ->  Conta quantas vezes cada coligada aparece para saber se é a última ocorrência.
_____________________________________________________________________________
Funções principais:
- *`encontrar_imagem()`* ->  Aguarda até que a imagem desejada apareça na tela com uma certa confiança, retornando suas coordenadas.
- *`primeira_vez()`* ->  Executa os cliques e inserções necessários ao acessar uma coligada pela primeira vez.
- *`nem_nem()`* ->  Processo reduzido para coligadas que já foram acessadas anteriormente.
_____________________________________________________________________________
Tecnologias utilizadas:
- `pyautogui` – automação de mouse e teclado
- `pyperclip` – manipulação da área de transferência (clipboard)
- `pandas` – leitura e manipulação da planilha
- `time.sleep` – controle de tempo entre ações
- `ImageNotFoundException` – tratamento de erro para imagem não encontrada

Observações:
- O script depende de uma pasta com imagens (`CAMINHO_IMAGENS`) e de um arquivo `.xlsx` com os dados.
- A automação é sensível à resolução e posição dos elementos na tela. Certifique-se de que as imagens de referência estão atualizadas.
- Os dados presentes neste repositório foram modificados e não correspondem aos dados reais da empresa. Foram utilizados apenas para fins de demonstração.
