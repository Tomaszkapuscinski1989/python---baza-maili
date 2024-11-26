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
import exel

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

        self.protocol("WM_DELETE_WINDOW", lambda x=self: contextManager.zamykanie(x))

        self.info = informacjeOprogramie.Info()
        self.mail = mail.Mail()
        self.exel = exel.Exel()

        menu = Menu(self)

        self.config(menu=menu)
        bazaDanych = Menu(
            menu, activebackground="black", activeforeground="white", tearoff=0
        )
        self.poczta = Menu(
            menu, activebackground="black", activeforeground="white", tearoff=0
        )

        self.dane = Menu(
            menu, activebackground="black", activeforeground="white", tearoff=0
        )

        self.wyszukaj = Menu(
            menu, activebackground="black", activeforeground="white", tearoff=0
        )

        self.informacje = Menu(
            menu, activebackground="black", activeforeground="white", tearoff=0
        )

        menu.add_cascade(label="Plik", menu=bazaDanych)
        bazaDanych.add_command(label="Nowy", command=self.nowaBaza)
        bazaDanych.add_command(label="Otwórz", command=self.otworzBaze)
        bazaDanych.add_separator()
        bazaDanych.add_command(
            label="Exit", command=lambda x=self: contextManager.zamykanie(x)
        )

        menu.add_cascade(label="Poczta", menu=self.poczta)
        self.poczta.add_command(
            label="Logowanie do poczty",
            command=self.mail.daneLogowania,
            state=DISABLED,
        )
        self.poczta.add_separator()

        self.poczta.add_command(
            label="wyślij wiadomość do wszystkich",
            command=self.mail.wyslijWiadomoscDoWszystkich,
            state=DISABLED,
        )

        menu.add_cascade(label="Dane", menu=self.dane)
        self.dane.add_command(
            label="Eksportuj do pliku exel",
            command=self.exel.export_do_exela,
            state=DISABLED,
        )
        self.dane.add_command(
            label="Importuj z pliku exel",
            command=self.exel.import_z_exela,
            state=DISABLED,
        )

        menu.add_cascade(label="Szukaj...", menu=self.wyszukaj)
        self.wyszukaj.add_command(
            label="według pola imię",
            command=lambda: self.szukaj("imie"),
            state=DISABLED,
        )
        self.wyszukaj.add_command(
            label="według pola nazwisko",
            command=lambda: self.szukaj("nazwisko"),
            state=DISABLED,
        )
        self.wyszukaj.add_command(
            label="według pola nazwa firmy",
            command=lambda: self.szukaj("nazaFirmy"),
            state=DISABLED,
        )

        menu.add_cascade(label="O programie", menu=self.informacje)
        self.informacje.add_command(
            label="Informacje o programie", command=self.info.info
        )

        ramka1 = Frame(self)
        ramka1.pack()

        self.listLabel = [
            ["Id:", 0, 0],
            ["Imię:", 1, 0],
            ["Nazwisko:", 2, 0],
            ["Telefon:", 3, 0],
            ["Nazwa Firmy:", 4, 0],
            ["Strona Internetowa:", 5, 0],
            ["Adres e-mail:", 6, 0],
        ]
        for item in self.listLabel:
            Label(ramka1, text=item[0]).grid(row=item[1], column=item[2])

        self.idEnter = Entry(ramka1)
        self.idEnter.grid(row=0, column=1, columnspan=3, sticky="WE")

        self.entryList = []

        entrys = [
            ["Imię", 1, 1, 3],
            ["Nazwisko", 2, 1, 3],
            ["111222333", 3, 1, 3],
            ["Nazwa firmy", 4, 1, 3],
            ["www.przykład.pl", 5, 1, 3],
            ["przykład@przykład.pl", 6, 1, 3],
        ]

        for item in entrys:
            teksPrzykład = item[0]
            t = Entry(ramka1, foreground="gray")
            t.grid(row=item[1], column=item[2], columnspan=item[3], sticky="WE")
            t.bind(
                "<FocusIn>",
                lambda event, x=teksPrzykład, y=t: contextManager.focusIn(x, y),
            )
            t.bind(
                "<FocusOut>",
                lambda event, x=teksPrzykład, y=t: contextManager.focusOut(x, y),
            )

            t.insert(0, teksPrzykład)
            self.entryList.append(t)

        self.stronaPrzycisk = ttk.Button(
            ramka1, text="Przedż do strony", state=DISABLED
        )
        self.stronaPrzycisk.grid(row=5, column=5)
        self.stronaPrzycisk.bind("<Button-1>", self.stronaKlienta)
        self.stronaPrzycisk.bind("<Return>", self.stronaKlienta)

        self.mailPrzycisk = ttk.Button(ramka1, text="Wyślij wiadomość", state=DISABLED)
        self.mailPrzycisk.grid(row=6, column=5)

        self.mailPrzycisk.bind(
            "<Button-1>",
            lambda event, y=self.entryList[5]: self.mail.wyslijWiadomosc(y),
        )
        self.mailPrzycisk.bind(
            "<Return>", lambda event, y=self.entryList[5]: self.mail.wyslijWiadomosc(y)
        )

        self.dodajKlientaPrzycisk = ttk.Button(
            ramka1, text="Dodaj klienta", state=DISABLED
        )
        self.dodajKlientaPrzycisk.grid(row=0, column=6)
        self.dodajKlientaPrzycisk.bind("<Button-1>", self.dodajKlienta)
        self.dodajKlientaPrzycisk.bind("<Return>", self.dodajKlienta)

        self.edytujKlientaPrzycisk = ttk.Button(
            ramka1, text="Edytuj klienta", state=DISABLED
        )
        self.edytujKlientaPrzycisk.grid(row=1, column=6)
        self.edytujKlientaPrzycisk.bind("<Button-1>", self.edytujKlienta)
        self.edytujKlientaPrzycisk.bind("<Return>", self.edytujKlienta)

        self.usunKlientaPrzycisk = ttk.Button(
            ramka1, text="Usuń klienta", state=DISABLED
        )
        self.usunKlientaPrzycisk.grid(row=2, column=6)
        self.usunKlientaPrzycisk.bind("<Button-1>", self.usunKlienta)
        self.usunKlientaPrzycisk.bind("<Return>", self.usunKlienta)

        self.tree_frame = Frame(self)
        self.tree_frame.pack(side=LEFT, fill=X, expand=True)

        tree_scrolly = ttk.Scrollbar(self.tree_frame)
        tree_scrolly.pack(side=RIGHT, fill=Y)
        tree_scrollx = ttk.Scrollbar(self.tree_frame, orient=HORIZONTAL)
        tree_scrollx.pack(side=BOTTOM, fill=X)
        self.refresh_button = ttk.Button(
            self.tree_frame,
            command=self.wyświetlWszystkichKlientow,
            text="Odśwież",
            state=DISABLED,
        )
        self.refresh_button.pack()

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

        myTreeHeading = [
            {"name": "#0", "width": 0, "stretch": False, "text": ""},
            {"name": "imie", "width": 140, "text": "Imię", "stretch": True},
            {
                "name": "nazwisko",
                "width": 40,
                "text": "Nazwisko",
                "stretch": True,
            },
            {
                "name": "telefon",
                "width": 40,
                "text": "Telefon",
                "stretch": True,
            },
            {
                "name": "nazwaFirmy",
                "width": 140,
                "text": "Nazwa firmy",
                "stretch": True,
            },
            {
                "name": "strona",
                "width": 40,
                "text": "Strona Internetowa",
                "stretch": True,
            },
            {
                "name": "mail",
                "width": 40,
                "text": "Adres e-mail",
                "stretch": True,
            },
            {"name": "oid", "width": 0, "stretch": False, "text": ""},
        ]

        for item in myTreeHeading:
            self.my_tree.column(
                item["name"], width=item["width"], stretch=item["stretch"], anchor=W
            )
            self.my_tree.heading(item["name"], text=item["text"], anchor=W)

        self.my_tree.tag_configure("odd", background=bg2, foreground="white")
        self.my_tree.tag_configure("even", background=bg3, foreground="white")

        self.my_tree.bind("<Double-Button-1>", self.wyswietlDaneKlienta)

        self.listaPol = [
            [self.idEnter, "Id:", "Id"],
            [self.entryList[0], "Imię:", "Imię"],
            [self.entryList[1], "Nazwisko", "Nazwisko"],
            [self.entryList[2], "Telefon", "111222333"],
            [self.entryList[3], "Nazwa firmy", "Nazwa firmy"],
            [self.entryList[4], "Strona internetoes", "www.przykład.pl"],
            [self.entryList[5], "Adres e-mail", "przykład@przykład.pl"],
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
            if self.entryList[4].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            else:
                messagebox.showinfo("info", "Strona zostanie otwarta")
                webbrowser.open(self.entryList[4].get())
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

                    c.execute(
                        """CREATE TABLE IF NOT EXISTS DaneLogowania (
                            port INTIGER DEFAULT 0,
                            serwer TEXT,
                            login TEXT,
                            haslo TEXT
                            ) """
                    )

                    c.execute(
                        """CREATE TABLE IF NOT EXISTS stopkaMaila(
                            zawartosc TEXT
                            ) """
                    )
            except:
                messagebox.showerror("Error", "Wystąpił błąd.")
            else:
                messagebox.showinfo("info", "Baza utworzona pomyślnie")
                self.wyświetlWszystkichKlientow()
                self.inicjuj()

    def otworzBaze(self):
        self.nazwaBazy = filedialog.askopenfilename(filetypes=(("db", "*.db"),))

        if self.nazwaBazy:
            nazwa = self.nazwaBazy.split("/")
            self.title(f"{self.nazwaProgramu} - {nazwa[-1]}")

            messagebox.showinfo("info", "Otwarcie bazy powiodło się.")
            self.wyświetlWszystkichKlientow()
            self.inicjuj()

    def inicjuj(self):
        self.poczta.entryconfigure(0, state="normal")
        self.poczta.entryconfigure(2, state="normal")
        self.dane.entryconfigure(0, state="normal")
        self.dane.entryconfigure(1, state="normal")
        self.wyszukaj.entryconfigure(0, state="normal")
        self.wyszukaj.entryconfigure(1, state="normal")
        self.wyszukaj.entryconfigure(2, state="normal")
        self.stronaPrzycisk.config(state="normal")
        self.mailPrzycisk.config(state="normal")

        self.dodajKlientaPrzycisk.config(state="normal")
        self.edytujKlientaPrzycisk.config(state="normal")
        self.usunKlientaPrzycisk.config(state="normal")
        self.refresh_button.config(state="normal")

        self.mail.nazwaBazy = self.nazwaBazy
        self.exel.nazwaBazy = self.nazwaBazy

    def dodajKlienta(self, e):

        if self.nazwaBazy:

            if self.entryList[0].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.entryList[1].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.entryList[2].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.entryList[3].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.entryList[4].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.entryList[5].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            else:
                try:
                    with contextManager.open_base(self.nazwaBazy) as c:
                        c.execute(
                            " INSERT INTO klienci VALUES (:imie, :nazwisko, :telefon, :nazaFirmy, :stronaInternetowa, :adresMail)",
                            {
                                "imie": self.entryList[0].get(),
                                "nazwisko": self.entryList[1].get(),
                                "telefon": int(self.entryList[2].get()),
                                "nazaFirmy": self.entryList[3].get(),
                                "stronaInternetowa": self.entryList[4].get(),
                                "adresMail": self.entryList[5].get(),
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
            if self.entryList[0].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.entryList[1].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.entryList[2].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.entryList[3].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.entryList[4].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            elif self.entryList[5].cget("foreground") == "gray":
                messagebox.showerror("Error", "Nie wybrano danych.")
            else:
                try:
                    with contextManager.open_base(self.nazwaBazy) as c:
                        c.execute(
                            " UPDATE klienci SET imie=:imie, nazwisko=:nazwisko, telefon=:telefon, nazaFirmy=:nazaFirmy, stronaInternetowa=:stronaInternetowa, adresMail=:adresMail WHERE oid = :k4",
                            {
                                "imie": self.entryList[0].get(),
                                "nazwisko": self.entryList[1].get(),
                                "telefon": int(self.entryList[2].get()),
                                "nazaFirmy": self.entryList[3].get(),
                                "stronaInternetowa": self.entryList[4].get(),
                                "adresMail": self.entryList[5].get(),
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
                        with contextManager.open_base(self.nazwaBazy) as c:
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

            for counter, value in enumerate(r, start=1):
                if counter % 2 == 0:
                    mode = "odd"
                else:
                    mode = "even"

                self.my_tree.insert(
                    parent="",
                    index="end",
                    iid=counter,
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
                    tags=(mode,),
                )

    def szukaj(self, pole):
        self.oknoWyszukiwania = Toplevel()
        self.oknoWyszukiwania.title("Szukaj")

        with contextManager.open_base(self.nazwaBazy) as c:
            c.execute(f"SELECT {pole} FROM klienci")
            lista = c.fetchall()

        lista = list(set(lista))
        lista.insert(0, "Wybierz")

        header = (
            "Imię"
            if pole == "imie"
            else ("Nazwisko" if pole == "nazwisko" else "Nazwa firmy")
        )
        szukajLabel = Label(self.oknoWyszukiwania, text=f"Szukaj według pola {header}:")
        szukajLabel.pack()
        self.wybor_id = ttk.Combobox(self.oknoWyszukiwania, value=lista)
        self.wybor_id.current(0)
        self.wybor_id.pack()

        self.refresh_button = ttk.Button(
            self.oknoWyszukiwania,
            command=lambda: self.szukaj_2(pole),
            text="Wyszukaj",
        )
        self.refresh_button.pack()

    def szukaj_2(self, pole):
        with contextManager.open_base(self.nazwaBazy) as c:
            c.execute(
                f"SELECT *, oid FROM klienci WHERE {pole} = :oid",
                {"oid": self.wybor_id.get()},
            )
            r = c.fetchall()

        if r:
            self.my_tree.delete(*self.my_tree.get_children())
            self.oknoWyszukiwania.destroy()

            for counter, value in enumerate(r, start=1):
                if counter % 2 == 0:
                    mode = "odd"
                else:
                    mode = "even"

                self.my_tree.insert(
                    parent="",
                    index="end",
                    iid=counter,
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
                    tags=(mode,),
                )
        else:
            self.oknoWyszukiwania.destroy()
            messagebox.showinfo("info", "")


if __name__ == "__main__":
    main = Main()
    main.mainloop()
