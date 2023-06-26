import re
import time
import pandas as pd
import pyautogui
import webbrowser
import urllib.parse

print('Iniciando...')

# Prompt user to enter the Excel file path
excel_file_path = 'C:\\Envio de Cobranças.xlsm'
try:
    with pd.ExcelFile(excel_file_path) as xl:
        # Read data from the Excel file
        bd = pd.read_excel(xl, sheet_name='BD')
        msg = pd.read_excel(xl, sheet_name='Mensagens')['MENSAGENS'].dropna().tolist()
except FileNotFoundError:
    print('Arquivo não encontrado, verifique o caminho e tente novamente.')
    exit()

# Create a dictionary of clients
clients = {}
exclude = []
for index, row in bd.iterrows():
    codcli = row['COD_CLIENTE']
    nome = row['CLIENTE']
    atraso = row['ATRASO']
    telefone = str(row['TEL ATUALIZADO'])
    nome = re.sub(r'[^a-zA-Z\s]', '', nome)
    excluir = row['RENEGOCIADOS']
    if pd.notna(excluir) and isinstance(excluir, (float, int)):
        exclude.append(int(excluir))
    if codcli not in clients:
        clients[codcli] = {'CLIENTE': nome, 'ATRASO': atraso, 'TELEFONE': telefone}
    else:
        if clients[codcli]['CLIENTE'] is None:
            clients[codcli]['CLIENTE'] = nome
        if clients[codcli]['ATRASO'] is None:
            clients[codcli]['ATRASO'] = atraso
        if clients[codcli]['TELEFONE'] is None:
            clients[codcli]['TELEFONE'] = telefone

# Loop through each client and send a message on WhatsApp
webbrowser.open(f'https://google.com/')
time.sleep(3)
pyautogui.press('enter')
webbrowser.open(f'https://web.whatsapp.com/')
time.sleep(180)
pyautogui.hotkey('ctrl', 'w')
webbrowser.open('https://web.whatsapp.com/send?phone={own_number}&text=ola')
time.sleep(180)
for codcli in clients:
    if len(clients[codcli]['TELEFONE']) == 13:
        # Exclude clients in the `exclude` set
        if codcli not in exclude:
            pessoa = clients[codcli]['CLIENTE']
            numero = clients[codcli]['TELEFONE']
            atraso = clients[codcli]['ATRASO']
            if codcli not in exclude:
                if atraso < 8:
                    texto = msg[0].replace('[clie]', pessoa)
                    print(msg)
                elif atraso < 15:
                    texto = msg[1].replace('[clie]', pessoa)
                    print(msg)
                elif atraso < 21:
                    texto = msg[2].replace('[clie]', pessoa)
                    print(msg)
                else:
                    texto = msg[3].replace('[clie]', pessoa)
                    print(msg)
                msg_encoded = urllib.parse.quote(texto)
                print(msg_encoded)

                # Open WhatsApp Web and send the message
                webbrowser.open(f'https://web.whatsapp.com/send?phone={numero}&text={msg_encoded}')
                time.sleep(5)
                pyautogui.press('enter')
                time.sleep(10)
                pyautogui.hotkey('ctrl', 'w')
                time.sleep(10)
pyautogui.hotkey('ctrl', 'w')
