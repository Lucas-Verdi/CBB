import time
import tkinter as tk
from tkinter import *
from tkinter import filedialog
import xlwings
import pyautogui
import win32com.client as win32
from pyautogui import sleep
import sys
from threading import Thread


class Th(Thread):

    def __init__(self, num):
        Thread.__init__(self)
        self.num = num

    def run(self):

        #Criando janela para selecionar o arquivo
        root = tk.Tk()
        root.withdraw()
        arquivo = filedialog.askopenfilename()

        #Abrindo a planilha selecionada
        pastadetrabalho = xlwings.Book(arquivo)

        #Abre o Excel em tela cheia
        excel_window = pyautogui.getWindowsWithTitle("Excel")[0]
        excel_window.maximize()
        #xl = win32.gencache.EnsureDispatch('Excel.Application')
        #xl.Visible = True
        #xl.Workbooks.Open(arquivo)
        #xl.ActiveWindow.WindowState = win32.constants.xlMaximized

        #Selecionando a planilha
        planilha = pastadetrabalho.sheets["Planilha1"]

        #Lendo os dados da coluna A e E
        coluna_a = planilha.range("A1:A" + str(planilha.cells.last_cell.row)).value
        coluna_e = planilha.range("E1:E" + str(planilha.cells.last_cell.row)).value

        #Lista para armazenar as linhas onde ocorrem as alterações de data
        alteracoes_data = []
        linhas_mesma_data = []

        #inserindo nova linha
        numero_linha = 1
        planilha.api.Rows(numero_linha).Insert()
        linhatexto1 = 1
        linhatexto2 = 1
        valortexto1 = 'DATA'
        valortexto2 = 'VALOR'
        colunavalor1 = 'F'
        colunavalor2 = 'G'
        planilha.range('{}{}'.format(colunavalor1, linhatexto1)).value = valortexto1
        planilha.range('{}{}'.format(colunavalor2, linhatexto2)).value = valortexto2

        #Loop para encontrar as alterações de data
        for i in range(1, len(coluna_a)):
            if coluna_a[i] != coluna_a[i - 1]:
                alteracoes_data.append(i + 1)  # Adiciona a linha onde ocorreu a alteração de data

        for j in range(0, len(coluna_a)):
            if coluna_a[j] == coluna_a[j - 1]:
                linhas_mesma_data.append(j)

        #Loop para inserir os dados na planilha
        selecao = alteracoes_data[0]
        selecao2 = selecao - 1
        termo = linhas_mesma_data[1]

        for auto in alteracoes_data:
            auto -= 1
            celula_selecionada = planilha.range("A{}".format(auto + 1))
            celula_selecionada.select()

            pyautogui.moveTo(400, 0)
            pyautogui.click()
            pyautogui.press('right')
            pyautogui.press('right')
            pyautogui.press('right')
            pyautogui.press('right')
            pyautogui.press('right')
            pyautogui.typewrite('=A{}'.format(auto + 1))
            # pyautogui.typewrite('=PROCV(A{};A:A;1;0)'.format(auto))
            pyautogui.press('right')
            pyautogui.typewrite('=SOMA(E{}:E{})'.format(termo, auto + 1))
            pyautogui.press('right')
            termo = auto + 2

        #Copiando dados inseridos para a planilha 2
        faixa_origem = 'F:G'
        valores = planilha.range(faixa_origem).value
        nova_planilha = pastadetrabalho.sheets.add('Planilha 2')
        nova_planilha.range('A1').value = valores

        #Selecionar a faixa de células com os dados na nova planilha
        faixa_dados = nova_planilha.range('A1:B{}'.format(nova_planilha.cells.last_cell.row))

        #Definir o filtro para excluir as linhas vazias na coluna A
        faixa_dados.api.AutoFilter(Field=1, Criteria1="<>")

        #Obter a coluna A filtrada (coluna A sem as linhas vazias)
        coluna_a_filtrada = nova_planilha.range('A2').expand('down').value

        #Remover o filtro
        faixa_dados.api.AutoFilterMode = False

        #Deletar as linhas vazias na coluna A
        linhas_deletar = []
        for i, valor in enumerate(coluna_a_filtrada, start=2):
            if not valor:
                linhas_deletar.append(i)

        nova_planilha.range('A1').api.EntireRow.Range(
            f"{nova_planilha.cells(linhas_deletar[0], 1).address}:{nova_planilha.cells(linhas_deletar[-1], 1).address}").api.Delete()

def start():
    a = Th(1)
    a.start()

#INTERFACE
janela = Tk()
janela.title('CobrançasBB')
Label1 = Label(janela, text='Insira a planilha de cobranças:')
Label1.grid(column=0, row=0, padx=10, pady=10)
Botao1 = Button(janela, text='Inserir')
Botao1.bind("<Button>",  lambda e: start())
Botao1.grid(column=0, row=1, padx=10, pady=10)
janela.mainloop()
