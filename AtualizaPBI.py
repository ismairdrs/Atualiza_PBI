import time
import os
import sys
import win32gui
from pywinauto.application import Application
from pywinauto import Desktop


def main(workbook, workspace="My workspace"):

    time.sleep(2)
    WORKBOOK = workbook
    WORKSPACE = workspace
    INIT_WAIT = 30000
    REFRESH_TIMEOUT = 15

    print("----- Fecha Power BI em execução")
    PROCNAME = "PBIDesktop.exe"
    try:
        app = Application(backend='uia').connect(path=PROCNAME)
        win = app.window(title_re='.*Power BI Desktop')
        win.close()
        time.sleep(2)
        print('Power BI foi fechado!')
    except:
        pass

    filtrar_titulo = workbook.split('.')[0].split('\\')[-1]
    titulo = f'{filtrar_titulo} - Power BI Desktop'


    print("----- Iniciando Power BI")
    try:
        os.system('start "" "' + workbook + '"')
    except Exception:
        print(Exception)
    time.sleep(25)

    PROCNAME = "PBIDesktop.exe"

    # Connect pywinauto
    print("----- Encontrando o Power BI")
    try:
        app = Application(backend='uia').connect(path=PROCNAME)
        win = app.window(title_re='.*Power BI Desktop')
        top_windows = Desktop(backend="uia").windows()
        win.wait("enabled", timeout=25)

        for w in top_windows:
            pbi = f'{w.window_text()}'
            if pbi == titulo:
                a: str = f'{pbi}'
                hwnd = win32gui.FindWindow(None, (a))
                win32gui.MoveWindow(hwnd, 30, 30, 900, 900, True)
                win32gui.GetFocus()
                win.wait("enabled", timeout=10)
    except:
        sys.exit(1)

    time.sleep(5)

    print('----- Atualizando')

    win.click_input(coords=(450, 100))
    time.sleep(40)


    print("----- Salvando")
    try:
        win.click_input(coords=(17, 13))
        win.click_input(coords=(17, 13))
        win.click_input(coords=(17, 13))
        win.click_input(coords=(17, 13))
    except Exception as e:
        print(e)

    time.sleep(5)
    win.wait("enabled", timeout=10)


    print('----- Publicar na Web')
    win.click_input(coords=(780, 100))
    time.sleep(3)
    WORKSPACE = "Meu workspace"

    publish_dialog = win.child_window(auto_id="KoPublishToGroupDialog")
    publish_dialog.workspaceDataItem.click_input()
    publish_dialog.workspaceDataItem.child_window().click_input()
    # print_control_identifiers()
    publish_dialog.Select.click()
    try:
        win.Replace.wait('visible', timeout=10)
    except Exception:
        pass
    if win.Replace.exists():
        win.Replace.click_input()
    time.sleep(5)
    win["Got it"].wait('visible', timeout=30000)
    win["Got it"].click_input()

    PROCNAME = "PBIDesktop.exe"
    try:
        app = Application(backend='uia').connect(path=PROCNAME)
        win = app.window(title_re='.*Power BI Desktop')
        win.close()
        time.sleep(2)
        print('----- Power BI foi publicado e fechado!')
    except:
        pass


if __name__ == '__main__':
    try:
        main(r'C:\Users\Junior\Desktop\Covid.pbix')

        print('## Fim ##')
    except Exception as e:
        print(e)
        sys.exit(1)
