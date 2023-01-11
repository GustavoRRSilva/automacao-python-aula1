import pyautogui;
import time;
import pandas as pd;
import numpy;
import openpyxl;
#passo 1 - Entrar no sistema da empresa(link do drive)
pyautogui.PAUSE = 2; #Pausa entre cada linha
pyautogui.hotkey("alt", "tab") # dar alt+tab para ir pro brave
pyautogui.hotkey("ctrl", "t") # abrir uma nova aba
#usar pyautogui.copy("url que usar") quando a url tiver caracteres especiais
pyautogui.write("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga");#Escrever o link
pyautogui.press("enter") # aperta para ir para o link
time.sleep(3);

#passo 2 - Navegar até o local do relatario(entrar na pasta explorar)
# Descobre a posição do mouse print(pyautogui.position())
pyautogui.click(437, 331, clicks = 2 );

#passo 3 - exportar o relátorio => Fazer o download
pyautogui.click()
pyautogui.click(1655, 189);
pyautogui.click(1423, 684)
time.sleep(3)
pyautogui.press("enter");

#passo 4 - calcular os indicadores (faturamento e quantidade de produtos)
tabela = pd.read_excel(r"C:\Users\Gustavo Rodrigues\Downloads\Vendas - Dez.xlsx");
print(tabela)
faturamento = tabela["Valor Final"].sum();
quantidade = tabela["Quantidade"].sum();
print(faturamento);
print(quantidade);

#passo 5 - Enviar email
pyautogui.hotkey("ctrl","t")
pyautogui.write("https://gmail.com/")
pyautogui.press("enter")
pyautogui.sleep(3);
pyautogui.click(147,209)
pyautogui.sleep(2);
pyautogui.write("gstvrrds@gmail.com");
pyautogui.press("enter")
time.sleep(2)
pyautogui.press("tab")
pyautogui.write("Enviando faturamento e quantidades vendidas referente ao dia anterior")
time.sleep(3);
pyautogui.press("tab");
pyautogui.write("Prezados Senhores");
pyautogui.press("enter");
pyautogui.press("enter");
pyautogui.write("Faturamento de ontem:R$" + str(faturamento))
pyautogui.press("enter");
pyautogui.write("Quantidades vendidas:"+ str(quantidade))
pyautogui.press("enter")
pyautogui.press("enter")
pyautogui.write("Atenciosamente");
pyautogui.press("enter");
pyautogui.write("Gustavo Rodrigues Ramos Da Silva");
pyautogui.click(1377, 914)
time.sleep(3)
pyautogui.click(155, 68)
pyautogui.write(r"C:\Users\Gustavo Rodrigues\Downloads")
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(251, 528)
pyautogui.write("Vendas - Dez")
time.sleep(1);
pyautogui.press("enter")
time.sleep(1);
pyautogui.click(1244, 908)


