import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    # nome
    pyautogui.click(927,515,duration=1.5)
    pyautogui.write(linha[0].value)
    # produto
    pyautogui.click(936,542,duration=1.5)
    pyautogui.write(linha[1].value)
    # quantidade
    pyautogui.click(940,569,duration=1.5)
    pyautogui.write(str(linha[2].value))
    # categoria
    pyautogui.click(1075,590,duration=1.5)
    pyautogui.write(linha[3].value)
    pyautogui.click(1035,614,duration=1.5)
    pyautogui.click(874,622,duration=1.5)
    pyautogui.click(944,586,duration=1.5)