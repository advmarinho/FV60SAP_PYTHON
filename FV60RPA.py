import openpyxl
import pyautogui as abrirsite
from openpyxl import load_workbook
from openpyxl import Workbook

# pyinstaller --onefile --noconsole .\nome.py
# pyinstaller --onefile --console .\nome.py

print("\t |#############################################################| \t")
print("\t |                            _           _                    | \t")
print("\t |                           | |         | |                   | \t")
print("\t |                  _ __ ___ | |__   ___ | |_                  | \t")
print("\t |                 | '__/ _ \| '_ \ / _ \| __|                 | \t")
print("\t |                 | | | (_) | |_) | (_) | |_                  | \t")
print("\t |                 |_|  \___/|_.__/ \___/ \__|                 | \t")
print("\t |                                                             | \t")
print("\t |                                                             | \t")
print("\t |    **Robô encontra-se atuando na tarefa de fazer FV60**     | \t")
print("\t |   ENTRAR NO SAP COM LOGIN E SENHA E DEMAIS DADOS DA FV60    | \t")
print("\t |                         \      /                            | \t")
print("\t |                         (\____/)                            | \t")
print("\t |                          (_oo_)                             | \t")
print("\t |                           ([])           Oi, vamos começar. | \t")
print("\t |                          __||__    \)                       | \t")
print("\t |                       []/______\[] /                        | \t")
print("\t |                       / \______/ \/                         | \t")
print("\t |                      /    /__\                              | \t")
print("\t |                     (\   /____\                             | \t")
print("\t |#############################################################| \n\n\t")

print("\t |   --> Robô pronto para iniciar <--   | \n\n")
print("\t |   --> Entre no SAP com seu Login e Senha <--   | \n\n")
abrirsite.sleep(1)

#______________________________________________________________________

def fazendoFv60():
        nome_arquivo = "C:\\_RPA\\FV60.xlsx"
        print(nome_arquivo)
        print("\n\n\t |   --> Robô ENCONTROU esses dados para fazer FV60 <--                   \n")
        print('Siga as instruções  ')

        planilha_aberta = load_workbook(filename=nome_arquivo)
        sheet_selecionada = planilha_aberta['FV60']
        linhaWS = input("A partir da linha tal: ")

        for linha in range(int(linhaWS), len(sheet_selecionada['A']) + 1):

            fornecedorFv = sheet_selecionada['A%s' % linha].value
            dataFv = sheet_selecionada['B%s' % linha].value
            valorFv = sheet_selecionada['C%s' % linha].value
            nomeFv = sheet_selecionada['D%s' % linha].value
            contaFv = sheet_selecionada['E%s' % linha].value
            


            fornecedorFv60 = str(fornecedorFv)
            dataFv60 = str(dataFv)
            valorFv60 = str(valorFv).replace('.', ',')
            nomeFv60 = str(nomeFv)
            contaFv60 = str(contaFv)

            var = input("\n Havendo uma ou mais FV60 para fazer, TECLE ENTER=(Colocando os dados) se não, TECLE 2=(sair):  ").lower()
            if var == input(' Pressione ENTER 2x \n'):

                fornecedorFv60 = (str(fornecedorFv))
                dataFv60 = (str(dataFv))
                valorFv60 = (str(valorFv).replace('.', ','))
                nomeFv60 = (str(nomeFv))
                contaFv60 = (str(contaFv))

                abrirsite.sleep(2)
                abrirsite.click(x=700, y=500)
                abrirsite.sleep(2)

                abrirsite.moveTo(x=700, y=500)
                abrirsite.sleep(1)
                abrirsite.click(x=700, y=500)
                abrirsite.sleep(1)
                abrirsite.press('tab', presses=2)
                abrirsite.sleep(1)
                abrirsite.write('FV60')
                abrirsite.sleep(1)
                abrirsite.press('enter')
                abrirsite.sleep(1)

                abrirsite.write(fornecedorFv60)
                abrirsite.sleep(1)
                abrirsite.press('tab')
                abrirsite.sleep(1)
                abrirsite.press('tab')
                abrirsite.sleep(1)
                abrirsite.write(dataFv60)
                abrirsite.sleep(1)
                abrirsite.press('tab')
                abrirsite.sleep(1)
                abrirsite.write('60003')
                abrirsite.sleep(1)
                abrirsite.press('tab')
                abrirsite.sleep(1)
                abrirsite.press('tab')
                abrirsite.sleep(1)
                abrirsite.press('tab')
                abrirsite.sleep(1)
                abrirsite.write(valorFv60)
                abrirsite.sleep(1)
                abrirsite.press('tab', presses=3)
                abrirsite.sleep(1)
                abrirsite.write('0001')
                abrirsite.sleep(1)
                abrirsite.press('tab')
                abrirsite.sleep(1)
                abrirsite.write(nomeFv60)
                abrirsite.sleep(1)
                abrirsite.press('tab', presses=2)
                abrirsite.sleep(1)
                abrirsite.write(contaFv60)
                abrirsite.sleep(1)
                abrirsite.press('tab', presses=2)
                abrirsite.sleep(1)
                abrirsite.write(valorFv60)
                abrirsite.sleep(1)
                abrirsite.press('tab', presses=6)
                abrirsite.sleep(1)
                abrirsite.write(nomeFv60)
                abrirsite.sleep(1)
                with abrirsite.hold('shift'):
                    abrirsite.press('tab', presses=21)
                abrirsite.sleep(1)
                abrirsite.press('right')
                abrirsite.sleep(1)
                abrirsite.press('enter')
                abrirsite.sleep(2)

            if var == '2':
                print("\t |#############################################################| \t")
                print("\t |                         \      /                            | \t")
                print("\t |                         (\____/)                            | \t")
                print("\t |                          (_oo_)         Obrigado.           | \t")
                print("\t |                           ([])            Tchau!            | \t")
                print("\t |                    (/    __||__    \)                       | \t")
                print("\t |                     \ []/______\[] /                        | \t")
                print("\t |                      \/ \______/ \/                         | \t")
                print("\t |                           /__\                              | \t")
                print("\t |                          /____\                             | \t")
                print("\t |#############################################################| \n\n\t")
                print(    "\n\t |   --> TERMINAMOS, ATÉ LOGO, TCHAU <--                   \n\t")
                exit()
while True:
    fazendoFv60()
    print('FIM')
