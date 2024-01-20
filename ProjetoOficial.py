import datetime
import pandas as pd
import webbrowser
from time import sleep
import pyautogui as pa
from tkinter import *

def aplicacao():
    df = pd.read_excel('Automação wpp 1.xlsx')

    today = datetime.date.today()
    year = today.year
    for i in range(0, len(df)):
        nome = df['Nome'][i]
        dia = df['Dia Aniversario'][i]
        mes = df['Mes Aniversario'][i]
        telefone = df['Telefone'][i]
        birthdate = datetime.date(year, mes, dia)

        if birthdate == today:
            msg = f'Olá {nome}, passando para te desejar um ótimo aniversário'
            link = f'https://web.whatsapp.com/send?phone={telefone}&text={msg}'

            webbrowser.open(link)
            sleep(10)

            pa.moveTo(x=1840, y=953)
            sleep(3)
            pa.click()

            texto = f'{nome} está de aniversário hoje'
            texto_aniversarios['text'] = texto
        else:
            print(f'{nome} não está de aniversário hoje!')

    # Point(x=1840, y=953)
            
def cancelar():
    quit()

janela = Tk()
janela.title('Automação Whatsapp')
janela.configure(background='#161616')
janela.geometry('350x250')
janela.resizable(True, True)
janela.maxsize(width=500, height=400)
janela.minsize(width=300, height=200)

texto_orientacao1 = Label(janela, text='Clique em "Iniciar"', bg='#161616', fg='white', 
                         font=("verdana", 10, 'bold'))
texto_orientacao2 = Label(janela, text=' para a Automação começar!', bg='#161616', fg='white', 
                         font=("verdana", 10, 'bold'))
texto_orientacao1.place(relx=0.25, rely=0.15)
texto_orientacao2.place(relx=0.10, rely=0.25)

botao_iniciar = Button(janela, text='Iniciar', command=aplicacao, bg='#a0a0a0')
botao_iniciar.place(relx=0.3, rely=0.55, relwidth=0.2, relheight=0.15)

botao_cancelar = Button(janela, text='Cancelar', command=cancelar, bg='#a0a0a0')
botao_cancelar.place(relx=0.55, rely=0.55, relwidth=0.2, relheight=0.15)


texto_aniversarios = Label(janela, text='', bg='#161616', fg='white', 
                            font=("verdana", 10, 'bold'))
texto_aniversarios.place(relx=0.08, rely=0.85)

janela.mainloop()