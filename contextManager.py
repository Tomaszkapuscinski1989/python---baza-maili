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


@contextmanager
def open_base(name):
    try:
        conn = sq.connect(name)
        c = conn.cursor()
        yield c
    finally:
        conn.commit()
        conn.close()


class BrakWartosciError(Exception):
    pass


def focusIn(x, y):
    if y.get() == x:
        y.delete(0, END)
        y.configure(foreground="white")


def focusOut(x, y):
    if y.get() == "":
        y.insert(0, x)
        y.configure(foreground="gray")


def zamykanie(rodzic):
    zamknij = messagebox.askquestion("askquestion", "Zamknąć program?")
    if zamknij == "yes":
        rodzic.quit()
