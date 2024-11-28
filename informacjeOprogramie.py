from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
import webbrowser


class Info:

    def info(self):
        oknoInformacji = Toplevel()
        oknoInformacji.geometry("200x200")
        oknoInformacji.title("O programie")

        wersja = Label(oknoInformacji, text="Wersja: 0.9 Beta")
        wersja.pack()
        autor = Label(oknoInformacji, text="Autor: Tomasz Kapuściński")
        autor.pack()
        kontaktAutor = ttk.Button(oknoInformacji, text="Kontakt z autorem")
        kontaktAutor.pack()

        kontaktAutor.bind("<Button-1>", self.stronaAutora)
        kontaktAutor.bind("<Return>", self.stronaAutora)

    def stronaAutora(self, e):
        webbrowser.open("https://tomaszkapuscinski1989.github.io")
