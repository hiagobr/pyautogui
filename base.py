import pyautogui
import pandas as pd
import subprocess
import time
import win32gui
import win32con
import pyperclip
import csv

pyautogui.PAUSE = 0.3
pyautogui.FAILSAFE = True


def abrir_arquivo(nome):
    arquivo = pd.read_excel(nome)
    return arquivo


def tela_preco_rebate(_usuario_, _senha1_, _senha2_):

    # Abrir o GAN
    command = "C:\Program Files (x86)\PuTTY\putty.exe -load GAN -l hiagobr"
    sp = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE)

    # Tab para ativar a janela
    pyautogui.press('tab')
    pyautogui.press('tab')

    # Maximizar GAN
    hwnd = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

    # Reconhecimento imagem login 1
    login = pyautogui.locateOnScreen('login1.png', region=(0, 692, 114, 40))
    count = 1
    while login is None:
        login = pyautogui.locateOnScreen('login1.png', region=(0, 692, 114, 40))
        print "tela primeiro login", count, login
        count += 1
    # Fim

    pyautogui.typewrite(_usuario_)
    pyautogui.keyDown('enter')
    pyautogui.typewrite(_senha1_)
    pyautogui.keyDown('enter')
    pyautogui.typewrite('1')
    pyautogui.keyDown('enter')

    # Reconhecimento imagem login 2
    login2 = pyautogui.locateOnScreen('login2.png', region=(57, 313, 284, 167))
    count = 1
    while login2 is None:
        login2 = pyautogui.locateOnScreen('login2.png', region=(57, 313, 284, 167))
        print "tela segundo login", count, login2
        count += 1
    # Fim

    pyautogui.typewrite(_senha2_)
    pyautogui.keyDown('enter')
    pyautogui.keyDown('enter')
    pyautogui.keyDown('P')


def precificao():
    pyautogui.keyDown('F')
    pyautogui.keyDown('enter')
    pyautogui.keyDown('enter')


def tela_precos(_sku_, _estado_, _rua_, _copia_p_q_m_):

    # Reconhecimento imagem login 2
    tela_sku = pyautogui.locateOnScreen('tela_sku.png', region=(14, 38, 297, 674))
    count = 1
    while tela_sku is None:
        tela_sku = pyautogui.locateOnScreen('tela_sku.png', region=(14, 38, 297, 674))
        print "tela_sku", count, tela_sku
        count += 1
    # Fim

    pyautogui.keyDown('enter')
    pyautogui.keyDown('esc')
    pyautogui.keyDown('enter')
    pyautogui.typewrite(_sku_)
    pyautogui.keyDown('enter')

    # Selecionar _estado_
    if _estado_ == "T":
        pass
    elif _estado_ == "SP":
        pyautogui.press('down')
    elif _estado_ == "RJ":
        pyautogui.press('down', 2)
    elif _estado_ == "PR":
        pyautogui.press('down', 3)
    elif _estado_ == "MG":
        pyautogui.press('down', 4)
    elif _estado_ == "BA":
        pyautogui.press('down', 5)
    elif _estado_ == "RS":
        pyautogui.press('down', 6)
    elif _estado_ == "DF":
        pyautogui.press('down', 7)
    elif _estado_ == "GO":
        pyautogui.press('down', 8)
    elif _estado_ == "PE":
        pyautogui.press('down', 9)

    pyautogui.keyDown('enter')

    if _rua_ == "T":
        pyautogui.keyDown('enter')
        pyautogui.typewrite(_copia_p_q_m_)
        pyautogui.keyDown('enter')
    elif _rua_ == "P":
        pyautogui.press('down')
        pyautogui.keyDown('enter')
    elif _rua_ == "M":
        pyautogui.press('down', 2)
        pyautogui.keyDown('enter')
    elif _rua_ == "Q":
        pyautogui.press('down', 3)
        pyautogui.keyDown('enter')

    # Reconhecimento imagem cadastro
    tela_cadastro = pyautogui.locateOnScreen('cadastro.png', region=(41, 518, 219, 66))
    time.sleep(1)
    if tela_cadastro is not None:
        pyautogui.typewrite(_sku_)
        pyautogui.keyDown('enter')
        pyautogui.keyDown('enter')

        pyautogui.keyDown('enter')
        pyautogui.keyDown('esc')
        pyautogui.keyDown('enter')
        pyautogui.typewrite(_sku_)
        pyautogui.keyDown('enter')

        # Selecionar _estado_
        if _estado_ == "T":
            pass
        elif _estado_ == "SP":
            pyautogui.press('down')
        elif _estado_ == "RJ":
            pyautogui.press('down', 2)
        elif _estado_ == "PR":
            pyautogui.press('down', 3)
        elif _estado_ == "MG":
            pyautogui.press('down', 4)
        elif _estado_ == "BA":
            pyautogui.press('down', 5)
        elif _estado_ == "RS":
            pyautogui.press('down', 6)
        elif _estado_ == "DF":
            pyautogui.press('down', 7)
        elif _estado_ == "GO":
            pyautogui.press('down', 8)
        elif _estado_ == "PE":
            pyautogui.press('down', 9)

        pyautogui.keyDown('enter')

        if _rua_ == "T":
            pyautogui.keyDown('enter')
            pyautogui.typewrite(_copia_p_q_m_)
            pyautogui.keyDown('enter')
        elif _rua_ == "P":
            pyautogui.press('down')
            pyautogui.keyDown('enter')
        elif _rua_ == "M":
            pyautogui.press('down', 2)
            pyautogui.keyDown('enter')
        elif _rua_ == "Q":
            pyautogui.press('down', 3)
            pyautogui.keyDown('enter')
    else:
        pass


def verificar_negativo():
    negativos = pyautogui.locateOnScreen('negativo.png', region=(802, 71, 243, 644))

    if negativos is not None:
        pyautogui.keyDown('enter')


def etiqueta(condicao, preco):
    # Precificar Etiqueta
    pyautogui.keyDown('Q')
    pyautogui.typewrite(condicao)
    pyautogui.keyDown('enter')
    pyautogui.typewrite(preco)
    pyautogui.keyDown('enter')
    pyautogui.keyDown('enter')
    time.sleep(1)

    verificar_negativo()

    time.sleep(1)
    pyautogui.click(1000, 308, 3)
    evidencia_etiqueta = pyperclip.paste()
    evidencias["etiqueta"] = evidencia_etiqueta


def promocao(condicao, preco, _inicial_, _final_):
    # Precificar Etiqueta
    pyautogui.keyDown('P')
    pyautogui.typewrite(condicao)
    pyautogui.keyDown('enter')
    pyautogui.typewrite(_inicial_)
    pyautogui.keyDown('enter')
    pyautogui.typewrite(_final_)
    pyautogui.keyDown('enter')
    pyautogui.typewrite(preco)
    pyautogui.keyDown('enter')
    time.sleep(1)

    verificar_negativo()

    pyautogui.click(209, 220, 3)
    evidencia_promo = pyperclip.paste()
    evidencias["promo"] = evidencia_promo


def zerar_promocao():
    # Precificar Etiqueta
    pyautogui.keyDown('P')
    pyautogui.typewrite('75')
    pyautogui.keyDown('enter')
    pyautogui.typewrite('311217')
    pyautogui.keyDown('enter')
    pyautogui.typewrite('311217')
    pyautogui.keyDown('enter')
    pyautogui.typewrite('0')
    pyautogui.keyDown('enter')
    pyautogui.click(209, 220, 3)
    evidencia_promo2 = pyperclip.paste()
    evidencias["promo2"] = evidencia_promo2


def zerar_fixar():
    pyautogui.keyDown('X')
    pyautogui.keyDown('0')
    pyautogui.keyDown('enter')
    pyautogui.keyDown('0')
    pyautogui.keyDown('enter')
    pyautogui.click(885, 278, 3)
    evidencia_fixar = pyperclip.paste()
    evidencias["fixar"] = evidencia_fixar


def fixar(valor1, valor2):
    pyautogui.keyDown('X')
    pyautogui.typewrite(valor1)
    pyautogui.keyDown('enter')
    pyautogui.typewrite(valor2)
    pyautogui.keyDown('enter')

    verificar_negativo()

    pyautogui.click(885, 278, 3)
    evidencia_fixar = pyperclip.paste()
    evidencias["fixar"] = evidencia_fixar

if __name__ == "__main__":

    # Abrir arquivo
    df = abrir_arquivo("template_preco_backup.xlsx")

    # Abrir GAN e entrar na tela de precificacao e rebate
    usuario = "hiagobr"
    senha1 = "fast@2017"
    senha2 = "3301"
    tela_preco_rebate(usuario, senha1, senha2)

    # Entrar na tela de precificacao
    precificao()

    # Dicionario de evidencias
    evidencias = {}

    # Variaveis
    for index, row in df.iterrows():
        # Variaveis
        sku = row['sku']
        estado = row['estado']
        rua = row['rua']
        copia_p_q_m = row['copia_rua_p_q_m']
        etiqueta_condicao = str(row['etiqueta_condicao'])
        etiqueta_valor = str(row['etiqueta_valor'])
        promo_condicao = str(row['promo_condicao'])
        promo_valor = str(row['promo_valor'])
        inicial = str(row['inicial']).replace("i-", "")
        final = str(row['inicial']).replace("f-", "")
        fixar1 = str(row['fixar1'])
        fixar2 = str(row['fixar2'])
        print sku, estado, rua, copia_p_q_m, etiqueta_condicao, etiqueta_valor, promo_condicao, promo_valor,
        print inicial, final

        # Funcoes
        tela_precos(sku, estado, rua, copia_p_q_m)
        time.sleep(1)

        fixar(fixar1, fixar2)
        time.sleep(1)

        etiqueta(etiqueta_condicao, etiqueta_valor)
        time.sleep(1)

        zerar_promocao()
        time.sleep(1)

        pyautogui.keyDown('esc')

    with open('evidencias.  csv', 'wb') as f:
        w = csv.writer(f)
        w.writerow(evidencias.keys())

