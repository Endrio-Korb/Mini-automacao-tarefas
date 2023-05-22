import win32api
import win32print
import os


if __name__ == "__main__":

    lista_impressoras = win32print.EnumPrinters(2)
    impressora = lista_impressoras[2]

    win32print.SetDefaultPrinter(impressora[2])

    caminho = r"caminho da pasta"

    lista_arquivos = os.listdir(caminho)

    for arquivo in lista_arquivos:
        win32api.ShellExecute(0, "print", arquivo, None, caminho, 0)