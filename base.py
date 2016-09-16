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
    pyautogui.keyDown('enter')

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

    # Reconhecimento imagem tela_sku
    tela_sku = pyautogui.locateOnScreen('tela_sku.png', region=(14, 38, 297, 674))
    count = 1
    while tela_sku is None:
        tela_sku = pyautogui.locateOnScreen('tela_sku.png', region=(14, 38, 297, 674))
        count += 1
    # Fim

    pyautogui.keyDown('enter')
    pyautogui.keyDown('esc')
    pyautogui.keyDown('enter')
    pyautogui.keyDown('enter')
    pyautogui.typewrite(_sku_)
    pyautogui.keyDown('enter')
    pyautogui.keyDown('enter')

    # Reconhecimento imagem tela_estado
    tela_estado = pyautogui.locateOnScreen('tela_estado.png', region=(59, 422, 227, 92))
    count = 1
    while tela_estado is None:
        tela_estado = pyautogui.locateOnScreen('tela_estado.png', region=(59, 422, 227, 92))
        count += 1
    # Fim

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

    # Reconhecimento imagem tela_rua
    tela_rua = pyautogui.locateOnScreen('tela_rua.png', region=(499, 318, 360, 63))
    count = 1
    while tela_rua is None:
        tela_rua = pyautogui.locateOnScreen('tela_rua.png', region=(499, 318, 360, 63))
        count += 1
    # Fim

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
        pyautogui.keyDown('enter')
        pyautogui.typewrite(_sku_)
        pyautogui.keyDown('enter')
        pyautogui.keyDown('enter')

        # Reconhecimento imagem tela_estado
        tela_estado = pyautogui.locateOnScreen('tela_estado.png', region=(59, 422, 227, 92))
        count = 1
        while tela_estado is None:
            tela_estado = pyautogui.locateOnScreen('tela_estado.png', region=(59, 422, 227, 92))
            count += 1
        # Fim

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

        # Reconhecimento imagem tela_rua
        tela_rua = pyautogui.locateOnScreen('tela_rua.png', region=(499, 318, 360, 63))
        count = 1
        while tela_rua is None:
            tela_rua = pyautogui.locateOnScreen('tela_rua.png', region=(499, 318, 360, 63))
            count += 1
        # Fim

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

    time.sleep(1)
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

    verificar_negativo()

    # Reconhecimento imagem login 2
    fim_etiqueta = pyautogui.locateOnScreen('fim_etiqueta.png', region=(226, 292, 94, 30))
    count = 1
    while fim_etiqueta is None:
        fim_etiqueta = pyautogui.locateOnScreen('fim_etiqueta.png', region=(226, 292, 94, 30))
        count += 1
    # Fim


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

    # Reconhecimento imagem login 2
    fim_promo = pyautogui.locateOnScreen('fim_promo.png', region=(85, 206, 107, 28))
    count = 1
    while fim_promo is None:
        fim_promo = pyautogui.locateOnScreen('fim_promo.png', region=(85, 206, 107, 28))
        count += 1
    # Fim


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

    # Reconhecimento imagem login 2
    fim_promo = pyautogui.locateOnScreen('fim_promo.png', region=(85, 206, 107, 28))
    count = 1
    while fim_promo is None:
        fim_promo = pyautogui.locateOnScreen('fim_promo.png', region=(85, 206, 107, 28))
        count += 1
        # Fim


def zerar_fixar():
    pyautogui.keyDown('X')
    pyautogui.keyDown('0')
    pyautogui.keyDown('enter')
    pyautogui.keyDown('0')
    pyautogui.keyDown('enter')

    # Reconhecimento imagem login 2
    fim_fixar = pyautogui.locateOnScreen('fim_fixar.png', region=(95, 259, 53, 39))
    count = 1
    while fim_fixar is None:
        fim_fixar = pyautogui.locateOnScreen('fim_fixar.png', region=(95, 259, 53, 39))
        count += 1
    # Fim


def fixar(valor1, valor2):
    pyautogui.keyDown('X')
    pyautogui.typewrite(valor1)
    pyautogui.keyDown('enter')
    pyautogui.typewrite(valor2)
    pyautogui.keyDown('enter')

    verificar_negativo()

    # Reconhecimento imagem login 2
    fim_fixar = pyautogui.locateOnScreen('fim_fixar.png', region=(95, 259, 53, 39))
    count = 1
    while fim_fixar is None:
        fim_fixar = pyautogui.locateOnScreen('fim_fixar.png', region=(95, 259, 53, 39))
        count += 1
        # Fim


def evidencias():

    # Promocao
    pyautogui.click(209, 220, 3)
    evidencia_promo_aux = pyperclip.paste()

    # Fixar
    pyautogui.click(885, 278, 3)
    evidencia_fixar_aux = pyperclip.paste()

    # Etiqueta
    pyautogui.click(1000, 308, 3)
    evidencia_etiqueta_aux = pyperclip.paste()

    return evidencia_promo_aux, evidencia_fixar_aux, evidencia_etiqueta_aux

if __name__ == "__main__":

    startTime = time.time()

    # Abrir arquivo
    df = abrir_arquivo("template_preco.xlsx")

    # Abrir GAN e entrar na tela de precificacao e rebate
    usuario = "hiagobr"
    senha1 = "fast@2017"
    senha2 = "3301"
    tela_preco_rebate(usuario, senha1, senha2)

    # Entrar na tela de precificacao
    precificao()

    # Dicionario de evidencias
    sku_teste = 'teste'
    fixar_teste = 'teste'
    promocao_teste = 'teste'
    etiqueta_teste = 'etiqueta'
    df_evidencias = pd.DataFrame([[sku_teste, fixar_teste, promocao_teste, etiqueta_teste]],
                                 columns=['sku', 'promocao', 'fixar', 'etiqueta'])

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
        inicial = str(row['inicial']).replace("i-", "   ")
        final = str(row['final']).replace("f-", "")
        fixar1 = str(row['fixar1'])
        fixar2 = str(row['fixar2'])
        print sku, estado, rua, copia_p_q_m, etiqueta_condicao, etiqueta_valor, promo_condicao, promo_valor,
        print inicial, final

        # Funcoes
        tela_precos(sku, estado, rua, copia_p_q_m)

        tela_sem_custo = pyautogui.locateOnScreen('custo_reposicao.png', region=(300, 316, 723, 67))
        time.sleep(1)

        if tela_sem_custo is None:
            zerar_fixar()
            time.sleep(0.5)

            # fixar(fixar1, fixar2)
            # time.sleep(0.5)

            etiqueta(etiqueta_condicao, etiqueta_valor)
            time.sleep(0.5)

            promocao(promo_condicao, promo_valor, inicial, final)
            time.sleep(0.5)

            # zerar_promocao()
            # time.sleep(0.5)

            # -----------------------------------Evidencias-----------------------------------
            evidencia_promo, evidencia_fixar, evidencia_etiqueta = evidencias()

            df_evidencias_temp = pd.DataFrame([[sku, evidencia_promo, evidencia_fixar, evidencia_etiqueta]],
                                              columns=['sku', 'promocao', 'fixar', 'etiqueta'])
            df_evidencias = df_evidencias.append(df_evidencias_temp, ignore_index=True)

            # Sair da tela de precificacao

            time.sleep(0.5)

            pyautogui.keyDown('esc')

        else:
            pyautogui.keyDown('enter')

            df_evidencias_temp = pd.DataFrame([[sku, 'sem custo de reposicao', 'sem custo de reposicao',
                                                's'
                                                'em custo de reposicao']],
                                              columns=['sku', 'promocao'
                                                              'o', 'fixar', 'etiqueta'])
            df_evidencias = df_evidencias.append(df_evidencias_temp, ignore_index=True)

    df_evidencias = df_evidencias[df_evidencias['sku'] != 'teste']
    df_evidencias.to_excel('evidencias.xlsx')

    endTime = time.time()

    tempo = endTime - startTime
    tempo_sku = tempo/len(df)

    print "Tempo total: ", tempo
    print tempo_sku























