from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import smtplib
from email.message import EmailMessage
from email.utils import formataddr

import contextManager


class Mail:
    def __init__(self):

        self.__port = ""
        self.__serwer = ""
        self.__login = ""
        self.__haslo = ""
        self.nazwaBazy = ""

    def wyslijWiadomosc(self, mail):
        self.mailEnter = mail
        if self.mailEnter.cget("foreground") == "gray":
            messagebox.showerror("Error", "Nie wybrano danych.")
        else:
            self.oknoWiadomości = Toplevel()

            self.oknoWiadomości.title("Wysyłanie widomości")

            wiadomoscRamka = Frame(self.oknoWiadomości)
            wiadomoscRamka.pack()

            wiadomoscLabel = Label(wiadomoscRamka, text="Wysyłanie wiaadomości")
            wiadomoscLabel.grid(row=0, column=0, columnspan=2)

            odLabel = Label(wiadomoscRamka, text="Nadawca:")
            odLabel.grid(row=1, column=0)
            self.odEnter = Entry(wiadomoscRamka)
            teksPrzykład = self.__login
            self.odEnter.insert(0, teksPrzykład)
            self.odEnter.grid(row=1, column=1)

            doLabel = Label(wiadomoscRamka, text="Odbiorca:")
            doLabel.grid(row=2, column=0)
            self.doEnter = Entry(wiadomoscRamka)
            self.doEnter.insert(0, self.mailEnter.get())
            self.doEnter.grid(row=2, column=1)

            jakoLabel = Label(wiadomoscRamka, text="Wyświetl jako:")
            jakoLabel.grid(row=3, column=0)
            self.jakoEnter = Entry(wiadomoscRamka, foreground="gray")
            teksPrzykład = self.__login
            self.jakoEnter.insert(0, teksPrzykład)
            self.jakoEnter.bind(
                "<FocusIn>",
                lambda event, x=teksPrzykład, y=self.jakoEnter: contextManager.focusIn(
                    x, y
                ),
            )
            self.jakoEnter.bind(
                "<FocusOut>",
                lambda event, x=teksPrzykład, y=self.jakoEnter: contextManager.focusOut(
                    x, y
                ),
            )

            self.jakoEnter.grid(row=3, column=1)

            tematLabel = Label(wiadomoscRamka, text="Temat:")
            tematLabel.grid(row=4, column=0)
            self.tematEnter = Entry(wiadomoscRamka, foreground="gray")
            teksPrzykład = "Brak tematu"
            self.tematEnter.insert(0, teksPrzykład)
            self.tematEnter.bind(
                "<FocusIn>",
                lambda event, x=teksPrzykład, y=self.tematEnter: contextManager.focusIn(
                    x, y
                ),
            )
            self.tematEnter.bind(
                "<FocusOut>",
                lambda event, x=teksPrzykład, y=self.tematEnter: contextManager.focusOut(
                    x, y
                ),
            )
            self.tematEnter.grid(row=4, column=1)

            trescLabel = Label(wiadomoscRamka, text="Wpisz treść wiadomości")
            trescLabel.grid(row=5, column=0, columnspan=2)

            self.trescEnter = Text(wiadomoscRamka)
            self.trescEnter.grid(row=6, column=0, columnspan=2)

            trescLabel2 = Label(wiadomoscRamka, text="Wpisz treść wiadomości[html5]")
            trescLabel2.grid(row=5, column=2, columnspan=2)

            self.trescEnter2 = Text(wiadomoscRamka)
            self.trescEnter2.grid(row=6, column=2, columnspan=2)

            kontaktAutor = ttk.Button(wiadomoscRamka, text="Wyślij")
            kontaktAutor.grid(row=7, column=1, columnspan=2)
            kontaktAutor.bind("<Button-1>", self.wysylanie)
            kontaktAutor.bind("<Return>", self.wysylanie)

    def wysylanie(self, e):
        try:
            msg = EmailMessage()
            msg["Subject"] = self.tematEnter.get()
            msg["From"] = formataddr((self.jakoEnter.get(), f"{self.__login}"))
            msg["To"] = self.mailEnter.get()
            msg["BCC"] = self.__login

            msg.set_content(self.trescEnter.get(1.0, END))
            msg.add_alternative(self.trescEnter2.get(1.0, END), subtype="html")

            with smtplib.SMTP(self.__serwer, self.__port) as server:
                server.starttls()
                server.login(self.__login, self.__haslo)
                server.sendmail(self.__login, self.mailEnter.get(), msg.as_string())

        except:
            messagebox.showerror("Błąd", "Wystąpił błąd podczas wysyłania")
        else:
            self.oknoWiadomości.destroy()
            messagebox.showinfo("info", "Wysłano wiadomość")

    def wyslijWiadomoscDoWszystkich(self):
        self.oknoWiadomości1 = Toplevel()

        self.oknoWiadomości1.title("Wysyłanie widomości do Wszystkich z bazy")

        wiadomoscRamka = Frame(self.oknoWiadomości1)
        wiadomoscRamka.pack()

        wiadomoscLabel = Label(wiadomoscRamka, text="Wysyłanie wiaadomości")
        wiadomoscLabel.grid(row=0, column=0, columnspan=2)

        odLabel = Label(wiadomoscRamka, text="Nadawca:")
        odLabel.grid(row=1, column=0)
        self.odEnter = Entry(wiadomoscRamka)
        teksPrzykład = self.__login
        self.odEnter.insert(0, teksPrzykład)
        self.odEnter.grid(row=1, column=1)

        jakoLabel = Label(wiadomoscRamka, text="Wyświetl jako:")
        jakoLabel.grid(row=3, column=0)
        self.jakoEnter = Entry(wiadomoscRamka, foreground="gray")
        teksPrzykład = self.__login
        self.jakoEnter.insert(0, teksPrzykład)
        self.jakoEnter.bind(
            "<FocusIn>",
            lambda event, x=teksPrzykład, y=self.jakoEnter: contextManager.focusIn(
                x, y
            ),
        )
        self.jakoEnter.bind(
            "<FocusOut>",
            lambda event, x=teksPrzykład, y=self.jakoEnter: contextManager.focusOut(
                x, y
            ),
        )

        self.jakoEnter.grid(row=3, column=1)

        tematLabel = Label(wiadomoscRamka, text="Temat:")
        tematLabel.grid(row=4, column=0)
        self.tematEnter = Entry(wiadomoscRamka, foreground="gray")
        teksPrzykład = "Brak tematu"
        self.tematEnter.insert(0, teksPrzykład)
        self.tematEnter.bind(
            "<FocusIn>",
            lambda event, x=teksPrzykład, y=self.tematEnter: contextManager.focusIn(
                x, y
            ),
        )
        self.tematEnter.bind(
            "<FocusOut>",
            lambda event, x=teksPrzykład, y=self.tematEnter: contextManager.focusOut(
                x, y
            ),
        )
        self.tematEnter.grid(row=4, column=1)

        trescLabel = Label(wiadomoscRamka, text="Wpisz treść wiadomości")
        trescLabel.grid(row=5, column=0, columnspan=2)

        self.trescEnter = Text(wiadomoscRamka)
        self.trescEnter.grid(row=6, column=0, columnspan=2)

        trescLabel2 = Label(wiadomoscRamka, text="Wpisz treść wiadomości[html5]")
        trescLabel2.grid(row=5, column=2, columnspan=2)

        self.trescEnter2 = Text(wiadomoscRamka)
        self.trescEnter2.grid(row=6, column=2, columnspan=2)

        kontaktAutor = ttk.Button(wiadomoscRamka, text="Wyślij")
        kontaktAutor.grid(row=7, column=1, columnspan=2)
        kontaktAutor.bind("<Button-1>", self.wysylanieDoWszystkich)
        kontaktAutor.bind("<Return>", self.wysylanieDoWszystkich)

    def wysylanieDoWszystkich(self, e):
        with contextManager.open_base(self.nazwaBazy) as c:
            c.execute("select adresMail FROM klienci")
            r = c.fetchall()
        try:
            for adresOdbiorcy in r:
                try:
                    msg = EmailMessage()
                    msg["Subject"] = self.tematEnter.get()
                    msg["From"] = formataddr((self.jakoEnter.get(), f"{self.__login}"))
                    msg["To"] = adresOdbiorcy[0]
                    msg["BCC"] = self.__login

                    msg.set_content(self.trescEnter.get(1.0, END))
                    msg.add_alternative(self.trescEnter2.get(1.0, END), subtype="html")

                    with smtplib.SMTP(self.__serwer, self.__port) as server:
                        server.starttls()
                        server.login(self.__login, self.__haslo)
                        server.sendmail(self.__login, adresOdbiorcy[0], msg.as_string())
                except:
                    messagebox.showerror(
                        "Błąd", f"Wystąpił błąd dla adresu {adresOdbiorcy[0]}"
                    )
        except:
            messagebox.showerror("Błąd", "Wystąpił błąd podczas wysyłania")
        else:
            self.oknoWiadomości1.destroy()
            messagebox.showinfo("info", "Wysłano wiadomości")

    def daneLogowania(self):

        self.oknoLogowania = Toplevel()
        self.oknoLogowania.geometry("400x300")
        self.oknoLogowania.title("Logowanie do poczty")

        logowanieRamka = Frame(self.oknoLogowania)
        logowanieRamka.pack()
        logowanieLabel = Label(logowanieRamka, text="Dane logowania do poczty")
        logowanieLabel.grid(row=0, column=0, columnspan=2)

        portLabel = Label(logowanieRamka, text="Port usługi smtp:")
        self.portEnter = Entry(logowanieRamka, foreground="gray")
        teksPrzykład = "Adres serwera"
        self.portEnter.insert(0, teksPrzykład)
        self.portEnter.bind(
            "<FocusIn>",
            lambda event, x=teksPrzykład, y=self.portEnter: contextManager.focusIn(
                x, y
            ),
        )
        self.portEnter.bind(
            "<FocusOut>",
            lambda event, x=teksPrzykład, y=self.portEnter: contextManager.focusOut(
                x, y
            ),
        )
        portLabel.grid(row=1, column=0)
        self.portEnter.grid(row=1, column=1)

        serwerLabel = Label(logowanieRamka, text="Adres serwera smtp:")
        self.serwerEnter = Entry(logowanieRamka, foreground="gray")
        teksPrzykład = "Adres serwera"
        self.serwerEnter.insert(0, teksPrzykład)
        self.serwerEnter.bind(
            "<FocusIn>",
            lambda event, x=teksPrzykład, y=self.serwerEnter: contextManager.focusIn(
                x, y
            ),
        )
        self.serwerEnter.bind(
            "<FocusOut>",
            lambda event, x=teksPrzykład, y=self.serwerEnter: contextManager.focusOut(
                x, y
            ),
        )
        serwerLabel.grid(row=2, column=0)
        self.serwerEnter.grid(row=2, column=1)

        loginLabel = Label(logowanieRamka, text="Login:")
        self.loginEnter = Entry(logowanieRamka, foreground="gray")
        teksPrzykład = "Login"
        self.loginEnter.insert(0, teksPrzykład)
        self.loginEnter.bind(
            "<FocusIn>",
            lambda event, x=teksPrzykład, y=self.loginEnter: contextManager.focusIn(
                x, y
            ),
        )
        self.loginEnter.bind(
            "<FocusOut>",
            lambda event, x=teksPrzykład, y=self.loginEnter: contextManager.focusOut(
                x, y
            ),
        )
        loginLabel.grid(row=3, column=0)
        self.loginEnter.grid(row=3, column=1)

        hasłoLabel = Label(logowanieRamka, text="Hasło:")
        self.hasloEnter = Entry(logowanieRamka, foreground="gray")
        teksPrzykład = "Hasło"
        self.hasloEnter.insert(0, teksPrzykład)
        self.hasloEnter.bind(
            "<FocusIn>",
            lambda event, x=teksPrzykład, y=self.hasloEnter: contextManager.focusIn(
                x, y
            ),
        )
        self.hasloEnter.bind(
            "<FocusOut>",
            lambda event, x=teksPrzykład, y=self.hasloEnter: contextManager.focusOut(
                x, y
            ),
        )
        hasłoLabel.grid(row=4, column=0)
        self.hasloEnter.grid(row=4, column=1)

        kontaktAutor = ttk.Button(logowanieRamka, text="usun z bazy")
        kontaktAutor.grid(row=5, column=0)
        kontaktAutor.bind("<Button-1>", self.usunZBazy)
        kontaktAutor.bind("<Return>", self.usunZBazy)

        kontaktAutor = ttk.Button(logowanieRamka, text="Dodaj do bazy")
        kontaktAutor.grid(row=5, column=1)
        kontaktAutor.bind("<Button-1>", self.dodajDobazy)
        kontaktAutor.bind("<Return>", self.dodajDobazy)

        kontaktAutor = ttk.Button(logowanieRamka, text="Zatwierdź i zamknij")
        kontaktAutor.grid(row=6, column=0, columnspan=2)
        kontaktAutor.bind("<Button-1>", self.weryfikacjaDanychLogowania)
        kontaktAutor.bind("<Return>", self.weryfikacjaDanychLogowania)

        with contextManager.open_base(self.nazwaBazy) as c:
            c.execute("select * FROM DaneLogowania")
            r = c.fetchall()

        if r:
            self.portEnter.delete(0, END)
            self.portEnter.configure(foreground="white")
            self.portEnter.insert(0, r[0][0])

            self.serwerEnter.delete(0, END)
            self.serwerEnter.configure(foreground="white")
            self.serwerEnter.insert(0, r[0][1])

            self.loginEnter.delete(0, END)
            self.loginEnter.configure(foreground="white")
            self.loginEnter.insert(0, r[0][2])

            self.hasloEnter.delete(0, END)
            self.hasloEnter.configure(foreground="white")
            self.hasloEnter.insert(0, r[0][3])

            self.__port = int(self.portEnter.get())
            self.__serwer = self.serwerEnter.get()
            self.__login = self.loginEnter.get()
            self.__haslo = self.hasloEnter.get()

    def weryfikacjaDanychLogowania(self, e):
        try:
            if (
                self.portEnter.cget("foreground") == "gray"
                or self.portEnter.get() == ""
            ):
                raise contextManager.BrakWartosciError
            if (
                self.serwerEnter.cget("foreground") == "gray"
                or self.serwerEnter.get() == ""
            ):
                raise contextManager.BrakWartosciError
            if (
                self.loginEnter.cget("foreground") == "gray"
                or self.loginEnter.get() == ""
            ):
                raise contextManager.BrakWartosciError
            if (
                self.hasloEnter.cget("foreground") == "gray"
                or self.hasloEnter.get() == ""
            ):
                raise contextManager.BrakWartosciError

            self.__port = int(self.portEnter.get())
            self.__serwer = self.serwerEnter.get()
            self.__login = self.loginEnter.get()
            self.__haslo = self.hasloEnter.get()
        except ValueError:
            messagebox.showerror("Error", "Numer portu musi być liczbą.")
        except contextManager.BrakWartosciError:
            messagebox.showerror("Error", "Poprawnie uzupełnij wszystkie dane.")
        except Exception:
            messagebox.showerror("Error", "Wystąpił błąd.")
        else:
            self.oknoLogowania.destroy()

    def dodajDobazy(self, e):
        try:
            if (
                self.portEnter.cget("foreground") == "gray"
                or self.portEnter.get() == ""
            ):
                raise contextManager.BrakWartosciError
            if (
                self.serwerEnter.cget("foreground") == "gray"
                or self.serwerEnter.get() == ""
            ):
                raise contextManager.BrakWartosciError
            if (
                self.loginEnter.cget("foreground") == "gray"
                or self.loginEnter.get() == ""
            ):
                raise contextManager.BrakWartosciError
            if (
                self.hasloEnter.cget("foreground") == "gray"
                or self.hasloEnter.get() == ""
            ):
                raise contextManager.BrakWartosciError

            with contextManager.open_base(self.nazwaBazy) as c:
                c.execute(
                    " INSERT INTO DaneLogowania VALUES (:port, :serwer, :login, :haslo)",
                    {
                        "port": int(self.portEnter.get()),
                        "serwer": self.serwerEnter.get(),
                        "login": self.loginEnter.get(),
                        "haslo": self.hasloEnter.get(),
                    },
                )
        except ValueError:
            messagebox.showerror("Error", "Numer portu musi być liczbą.")
        except contextManager.BrakWartosciError:
            messagebox.showerror("Error", "Poprawnie uzupełnij wszystkie dane.")
        except Exception:
            messagebox.showerror("Error", "Wystąpił błąd.")
        else:
            self.oknoLogowania.destroy()

    def usunZBazy(self, e):
        try:
            with contextManager.open_base(self.nazwaBazy) as c:
                c.execute(
                    """DELETE FROM DaneLogowania  where oid=:oid""",
                    {"oid": 1},
                )
        except Exception:
            messagebox.showerror("Error", "Wystąpił błąd podczas usuwania wpisu.")
        else:
            messagebox.showinfo("Info", "Usunięto wpis.")
