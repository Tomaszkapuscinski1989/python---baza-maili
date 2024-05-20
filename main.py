from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
import webbrowser
from contextlib import contextmanager
import sqlite3 as sq
import os
import smtplib
from email.message import EmailMessage
from email.utils import formataddr

import informacjeOprogramie
import mail
import contextManager

fontDuzy = ("Times", 18)
fontMaly = ("Times", 14)
bg1 = "#3b3a38"
bg2 = "#66625f"
bg3 = "#302c28"
fg1 = "white"


class Main(Tk):

    def __init__(self, *args, **kwargs):
        Tk.__init__(self, *args, **kwargs)
        self.nazwaProgramu = "Baza klientów"
        self.title(self.nazwaProgramu)
        self.geometry("800x400")
        self.nazwaBazy = ""

        # self.idEnter = ""
        # self.imieEnter = ""
        # self.nazwiskoEnter = ""
        # self.telefonEnter = ""
        # self.nazwaEnter = ""
        # self.stronaEnter = ""
        # self.mailEnter = ""

        self.protocol("WM_DELETE_WINDOW", lambda x=self: contextManager.zamykanie(x))

        self.info = informacjeOprogramie.Info()
        self.mail = mail.Mail()

        menu = Menu(self)

        self.config(menu=menu)
        bazaDanych = Menu(
            menu, activebackground="black", activeforeground="white", tearoff=0
        )
        poczta = Menu(
            menu, activebackground="black", activeforeground="white", tearoff=0
        )

        informacje = Menu(
            menu, activebackground="black", activeforeground="white", tearoff=0
        )

        menu.add_cascade(label="Plik", menu=bazaDanych)
        bazaDanych.add_command(label="Nowy", command=self.nowaBaza)
        bazaDanych.add_command(label="Otwórz", command=self.otworzBaze)
        bazaDanych.add_separator()
        bazaDanych.add_command(
            label="Exit", command=lambda x=self: contextManager.zamykanie(x)
        )

        menu.add_cascade(label="Poczta", menu=poczta)
        poczta.add_command(
            label="Logowanie do poczty",
            command=self.mail.daneLogowania,
        )
        poczta.add_separator()
        poczta.add_command(label="Wyślij wiadomość")
        poczta.add_command(label="wyślij wiadomość do wszystkich")

        menu.add_cascade(label="O programie", menu=informacje)
        informacje.add_command(label="Informacje o programie", command=self.info.info)

        ramka1 = Frame(self)
        ramka1.pack()

        idLabel = Label(ramka1, text="Id:")
        self.idEnter = Entry(ramka1)
        idLabel.grid(row=0, column=0)
        self.idEnter.grid(row=0, column=1)

        imieLabel = Label(ramka1, text="Imię:")
        self.imieEnter = Entry(ramka1, foreground="gray")
        teksPrzykład = "Imię"
        self.imieEnter.insert(0, teksPrzykład)
        self.imieEnter.bind(
            "<FocusIn>",
            lambda event, x=teksPrzykład, y=self.imieEnter: contextManager.focusIn(
                x, y
            ),
        )
        self.imieEnter.bind(
            "<FocusOut>",
            lambda event, x=teksPrzykład, y=self.imieEnter: contextManager.focusOut(
                x, y
            ),
        )
        imieLabel.grid(row=1, column=0)
        self.imieEnter.grid(row=1, column=1)

        nazwiskoLabel = Label(ramka1, text="Nazwisko:")
        self.nazwiskoEnter = Entry(ramka1, foreground="gray")
        teksPrzykład = "Nazwisko"
        self.nazwiskoEnter.insert(0, teksPrzykład)
        self.nazwiskoEnter.bind(
            "<FocusIn>",
            lambda event, x=teksPrzykład, y=self.nazwiskoEnter: contextManager.focusIn(
                x, y
            ),
        )
        self.nazwiskoEnter.bind(
            "<FocusOut>",
            lambda event, x=teksPrzykład, y=self.nazwiskoEnter: contextManager.focusOut(
                x, y
            ),
        )
        nazwiskoLabel.grid(row=2, column=0)
        self.nazwiskoEnter.grid(row=2, column=1)

        telefonLabel = Label(ramka1, text="Telefon:")
        self.telefonEnter = Entry(ramka1, foreground="gray")
        teksPrzykład = "111222333"
        self.telefonEnter.insert(0, teksPrzykład)
        self.telefonEnter.bind(
            "<FocusIn>",
            lambda event, x=teksPrzykład, y=self.telefonEnter: contextManager.focusIn(
                x, y
            ),
        )
        self.telefonEnter.bind(
            "<FocusOut>",
            lambda event, x=teksPrzykład, y=self.telefonEnter: contextManager.focusOut(
                x, y
            ),
        )
        telefonLabel.grid(row=3, column=0)
        self.telefonEnter.grid(row=3, column=1)

        nazwaLabel = Label(ramka1, text="Nazwa firmy:")
        self.nazwaEnter = Entry(ramka1, foreground="gray")
        teksPrzykład = "Nazwa firmy"
        self.nazwaEnter.insert(0, teksPrzykład)
        self.nazwaEnter.bind(
            "<FocusIn>",
            lambda event, x=teksPrzykład, y=self.nazwaEnter: contextManager.focusIn(
                x, y
            ),
        )
        self.nazwaEnter.bind(
            "<FocusOut>",
            lambda event, x=teksPrzykład, y=self.nazwaEnter: contextManager.focusOut(
                x, y
            ),
        )
        nazwaLabel.grid(row=4, column=0)
        self.nazwaEnter.grid(row=4, column=1)

        stronaLabel = Label(ramka1, text="Strona internetowa:")
        self.stronaEnter = Entry(ramka1, foreground="gray")
        teksPrzykład = "www.przykład.pl"
        self.stronaEnter.insert(0, teksPrzykład)
        self.stronaEnter.bind(
            "<FocusIn>",
            lambda event, x=teksPrzykład, y=self.stronaEnter: contextManager.focusIn(
                x, y
            ),
        )
        self.stronaEnter.bind(
            "<FocusOut>",
            lambda event, x=teksPrzykład, y=self.stronaEnter: contextManager.focusOut(
                x, y
            ),
        )
        stronaLabel.grid(row=5, column=0)
        self.stronaEnter.grid(row=5, column=1)

        mailLabel = Label(ramka1, text="Adres e-mail:")
        self.mailEnter = Entry(ramka1, foreground="gray")
        teksPrzykład = "przykład@przykład.pl"
        self.mailEnter.insert(0, teksPrzykład)
        self.mailEnter.bind(
            "<FocusIn>",
            lambda event, x=teksPrzykład, y=self.mailEnter: contextManager.focusIn(
                x, y
            ),
        )
        self.mailEnter.bind(
            "<FocusOut>",
            lambda event, x=teksPrzykład, y=self.mailEnter: contextManager.focusOut(
                x, y
            ),
        )
        mailLabel.grid(row=6, column=0)
        self.mailEnter.grid(row=6, column=1)

        kontaktAutor = ttk.Button(ramka1, text="Przedż do strony")
        kontaktAutor.grid(row=5, column=2)
        kontaktAutor.bind("<Button-1>", self.stronaKlienta)
        kontaktAutor.bind("<Return>", self.stronaKlienta)

        kontaktAutor = ttk.Button(ramka1, text="Wyślij wiadomość")
        kontaktAutor.grid(row=6, column=2)

        kontaktAutor.bind(
            "<Button-1>",
            lambda event, y=self.mailEnter: self.mail.wyslijWiadomosc(y),
        )
        kontaktAutor.bind(
            "<Return>", lambda event, y=self.mailEnter: self.mail.wyslijWiadomosc(y)
        )

        kontaktAutor = ttk.Button(ramka1, text="Dodaj klienta")
        kontaktAutor.grid(row=0, column=3)
        kontaktAutor.bind("<Button-1>", self.dodajKlienta)
        kontaktAutor.bind("<Return>", self.dodajKlienta)

        kontaktAutor = ttk.Button(ramka1, text="Edytuj klienta")
        kontaktAutor.grid(row=1, column=3)
        kontaktAutor.bind("<Button-1>", self.edytujKlienta)
        kontaktAutor.bind("<Return>", self.edytujKlienta)

        kontaktAutor = ttk.Button(ramka1, text="Usuń klienta")
        kontaktAutor.grid(row=2, column=3)
        kontaktAutor.bind("<Button-1>", self.usunKlienta)
        kontaktAutor.bind("<Return>", self.usunKlienta)

        self.tree_frame = Frame(self)
        self.tree_frame.pack(side=LEFT, fill=X, expand=True)

        tree_scrolly = ttk.Scrollbar(self.tree_frame)
        tree_scrolly.pack(side=RIGHT, fill=Y)
        tree_scrollx = ttk.Scrollbar(self.tree_frame, orient=HORIZONTAL)
        tree_scrollx.pack(side=BOTTOM, fill=X)

        self.my_tree = ttk.Treeview(
            self.tree_frame,
            style="mystyle.Treeview",
            selectmode="browse",
            yscrollcommand=tree_scrolly.set,
            xscrollcommand=tree_scrollx.set,
        )
        self.my_tree.pack(fill=X, expand=True)

        tree_scrolly.config(command=self.my_tree.yview)
        tree_scrollx.config(command=self.my_tree.xview)

        self.my_tree["columns"] = (
            "imie",
            "nazwisko",
            "telefon",
            "nazwaFirmy",
            "strona",
            "mail",
            "oid",
        )

        self.my_tree.column("#0", width=0, stretch=NO)
        self.my_tree.heading("#0", text="", anchor=E)

        self.my_tree.column("imie", anchor=W, width=140)
        self.my_tree.heading("imie", text="Imię", anchor=W)

        self.my_tree.column("nazwisko", anchor=W, width=40)
        self.my_tree.heading("nazwisko", text="Nazwisko", anchor=W)

        self.my_tree.column("telefon", anchor=W, width=40)
        self.my_tree.heading("telefon", text="Telefon", anchor=W)

        self.my_tree.column("nazwaFirmy", anchor=W, width=140)
        self.my_tree.heading("nazwaFirmy", text="Nazwa firmy", anchor=W)

        self.my_tree.column("strona", anchor=W, width=40)
        self.my_tree.heading("strona", text="Strona Internetowa", anchor=W)

        self.my_tree.column("mail", anchor=W, width=40)
        self.my_tree.heading("mail", text="Adres e-mail", anchor=W)

        self.my_tree.column("oid", width=0, stretch=NO)
        self.my_tree.heading("oid", text="", anchor=E)

        self.my_tree.tag_configure("odd", background=bg2, foreground="white")
        self.my_tree.tag_configure("even", background=bg3, foreground="white")

        self.my_tree.bind("<Double-Button-1>", self.wyswietlDaneKlienta)

        self.listaPol = [
            [self.idEnter, "Id:", "Id"],
            [self.imieEnter, "Imię:", "Imię"],
            [self.nazwiskoEnter, "Nazwisko", "Nazwisko"],
            [self.telefonEnter, "Telefon", "111222333"],
            [self.nazwaEnter, "Nazwa firmy", "Nazwa firmy"],
            [self.stronaEnter, "Strona internetoes", "www.przykład.pl"],
            [self.mailEnter, "Adres e-mail", "przykład@przykład.pl"],
        ]

        self.wyświetlWszystkichKlientow()

    def wyswietlDaneKlienta(self, e):

        for i in self.listaPol:

            i[0].delete(0, END)

        selected = self.my_tree.focus()
        self.values = self.my_tree.item(selected, "values")

        for x, i in enumerate(self.listaPol, -1):
            i[0].configure(foreground="white")
            i[0].insert(0, self.values[x])

    def stronaKlienta(self, e):
        if self.nazwaBazy:

            if self.stronaEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            else:

                messagebox.showinfo("info", "Strona zostanie otwarta")
                webbrowser.open(self.stronaEnter.get())

        else:
            messagebox.showerror("Error", "Nie uruchomiono żadnej bazy")

    def nowaBaza(self):

        self.nazwaBazy = filedialog.asksaveasfilename(
            defaultextension="*.db", filetypes=(("db", "*.db"),)
        )

        if self.nazwaBazy:

            try:
                nazwa = self.nazwaBazy.split("/")
                if os.path.exists(f"{nazwa[-1]}"):
                    os.remove(f"{nazwa[-1]}")

                self.title(f"{self.nazwaProgramu}- {nazwa[-1]}")

                with contextManager.open_base(self.nazwaBazy) as c:
                    c.execute(
                        """CREATE TABLE IF NOT EXISTS klienci (
                            imie TEXT,
                            nazwisko TEXT,
                            telefon INTIGER DEFAULT 0,
                            nazaFirmy TEXT,
                            stronaInternetowa TEXT,
                            adresMail TEXT
                            ) """
                    )
            except:
                messagebox.showerror("Error", "Wystąpił błąd.")
            else:
                messagebox.showinfo("info", "Baza utworzona pomyślnie")

    def otworzBaze(self):

        self.nazwaBazy = filedialog.askopenfilename(filetypes=(("db", "*.db"),))

        if self.nazwaBazy:

            nazwa = self.nazwaBazy.split("/")
            self.title(f"{self.nazwaProgramu} - {nazwa[-1]}")

            messagebox.showinfo("info", "Otwarcie bazy powiodło się.")
            self.wyświetlWszystkichKlientow()

    def dodajKlienta(self, e):
        if self.nazwaBazy:

            if self.imieEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.nazwiskoEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.telefonEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.nazwaEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.stronaEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.mailEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            else:
                try:
                    with contextManager.open_base(self.nazwaBazy) as c:
                        c.execute(
                            " INSERT INTO klienci VALUES (:imie, :nazwisko, :telefon, :nazaFirmy, :stronaInternetowa, :adresMail)",
                            {
                                "imie": self.imieEnter.get(),
                                "nazwisko": self.nazwiskoEnter.get(),
                                "telefon": int(self.telefonEnter.get()),
                                "nazaFirmy": self.nazwaEnter.get(),
                                "stronaInternetowa": self.stronaEnter.get(),
                                "adresMail": self.mailEnter.get(),
                            },
                        )
                except Exception:
                    messagebox.showerror("Error", "Wystąpił błąd.")

                else:
                    messagebox.showinfo("Info", "Pomyślnie dodano nowy wpis.")
                    self.wyświetlWszystkichKlientow()

        else:
            messagebox.showerror("Error", "Nie uruchomiono żadnej bazy")

    def edytujKlienta(self, e):
        if self.nazwaBazy:
            if self.imieEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.nazwiskoEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.telefonEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.nazwaEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.stronaEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.mailEnter.cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            else:
                try:
                    with contextManager.open_base(self.nazwaBazy) as c:
                        c.execute(
                            " UPDATE klienci SET imie=:imie, nazwisko=:nazwisko, telefon=:telefon, nazaFirmy=:nazaFirmy, stronaInternetowa=:stronaInternetowa, adresMail=:adresMail WHERE oid = :k4",
                            {
                                "imie": self.imieEnter.get(),
                                "nazwisko": self.nazwiskoEnter.get(),
                                "telefon": int(self.telefonEnter.get()),
                                "nazaFirmy": self.nazwaEnter.get(),
                                "stronaInternetowa": self.stronaEnter.get(),
                                "adresMail": self.mailEnter.get(),
                                "k4": self.idEnter.get(),
                            },
                        )
                except Exception:
                    messagebox.showerror("Error", "Wystąpił błąd podczas edytowania.")
                else:
                    messagebox.showinfo("Info", "Pomyślnie edytowano wpis.")
                    self.wyświetlWszystkichKlientow()
        else:
            messagebox.showerror("Error", "Nie uruchomiono żadnej bazy")

    def usunKlienta(self, e):
        if self.nazwaBazy:
            if self.idEnter.get():
                self.zamknij = messagebox.askquestion("askquestion", "Usunąć wpis?")
                if self.zamknij == "yes":
                    try:
                        with self.open_base(self.nazwaBazy) as c:
                            c.execute(
                                """DELETE FROM klienci where oid=:oid""",
                                {"oid": self.idEnter.get()},
                            )
                    except Exception:
                        messagebox.showerror(
                            "Error", "Wystąpił błąd podczas usuwania wpisu."
                        )
                    else:
                        messagebox.showinfo("Info", "Usunięto wpis.")
                        self.wyświetlWszystkichKlientow()

            else:
                messagebox.showerror("Error", "Nie wybrano danych")

        else:
            messagebox.showerror("Error", "Nie uruchomiono żadnej bazy")

    def wyświetlWszystkichKlientow(self):
        if self.nazwaBazy:
            self.my_tree.delete(*self.my_tree.get_children())

            with contextManager.open_base(self.nazwaBazy) as c:
                c.execute("select *, oid FROM klienci")
                r = c.fetchall()

            self.counter = 1
            for value in r:

                if self.counter % 2 == 0:

                    self.my_tree.insert(
                        parent="",
                        index="end",
                        iid=self.counter,
                        text="",
                        values=(
                            value[0],
                            value[1],
                            value[2],
                            value[3],
                            value[4],
                            value[5],
                            value[-1],
                        ),
                        tags=("odd",),
                    )
                else:
                    self.my_tree.insert(
                        parent="",
                        index="end",
                        iid=self.counter,
                        text="",
                        values=(
                            value[0],
                            value[1],
                            value[2],
                            value[3],
                            value[4],
                            value[5],
                            value[-1],
                        ),
                        tags=("even",),
                    )
                self.counter += 1


if __name__ == "__main__":

    main = Main()

    main.mainloop()
