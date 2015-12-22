# ! /usr/bin/env python
# -*- coding: utf-8 -*-
__author__ = 'Marcin Pieczyński'

"""
Tworzy interface do wprowadzania danych: grup leków oraz ich skłądu do filtrowania danych z plików excel.
Finalne dane zapisuje w bazie danych SQLite w pliku db.
"""

"""
DO ZROBIENIA :
NAZWA PROFILU i ostrzeżenia związane z nazwą !!!!!!!!!!!!!!!!!!

"""

import sqlite3 as lite
from edc_sqlite import SQliteEdit

import Tkinter
import tkMessageBox
import ttk
import sys
import tkFileDialog
import os

Men_info = "\t \t \t \tGarść informacji jak posługiwać się Menadżerem profili.\n\n" \
          "W celach pokazowych przygotowałem kilka przykładowych profili gotowych do wykorzystania. " \
          "W celu ich obejrzenia należy wybrać jeden z nich i kliknąć EDYTÓJ PROFIL. " \
          "Dalsze czytanie instrukcji zalecam po otworzeniu jednego z przykładowych profili. \n\n" \
          "Zapis profilu:\n" \
          "Aby utworzyć nowy profil należy wprowadzić jego nazwę oraz informacje dotyczące leków, " \
          "których dane chcemy filtrować. Tworząc własne profile do filtrowania danych należy ściśle " \
          "wzorować sie na profilach pokazowych. Nazwa profilu musi być nieprzerwanym ciągiem liter i " \
          "cyfr bez polskich znaków oraz spacji.\n\n" \
          "Edycja profilu:\n" \
          "Edycja profilu pozwala na modyfikowanie istniejących profili.\n\n"\
          "Usuwanie profilu:\n" \
          "Bezpowrotnie usuwa profil z komputera.\n\n"\
          "Dolna część panelu służy do wprowadzania szczegółowych danych w profilach.\n\n" \
          "Lista regionów \ cegieł:\n" \
          "Należy wprowadzić nazwy, odzielone przecinkami, poszczególnych regionów i cegieł jakie nas interesują. " \
          "Na początku należy wprowadzić nazwę regionu, np: 'A4', a następnie wprowadzić nazwy cegieł, " \
          "np: A4, A4.01, A4.02 itd.\n\n" \
          "" \
          "" \
          "Cześć lewa (kolor szary):\n" \
          "Należy wprowadzić nazwy dla poszczególnych zakładek/gup leków - nazwy grup staną się w pliku wyjściowym excel " \
          "nazwami zakładek. W poszczególnych grupach leków należy następnie wprowadzić nazwy leków oddzielone " \
          "od siebie przecinkami. Na podstawie wprowadzonych informacji program zbierze dane PnVa oraz UNITS " \
          "dla podanych leków w wprowadzonych cegłach i zapisze je w dwóch plikach wyjściowych excel, " \
          "osobno z danymi PnVa i danymi UNITS. Liczbę zakładek/grup leków można zmieniać przyciskami '+' i '-'.\n\n" \
          "" \
          "Cześć prawa (kolor biały):\n" \
          "Należy wprowadzić nazwy dla poszczególnych leków oddzielone przecinkami." \
          "Program dla wprowadzonych leków zbierze dane PnVa oraz UNITS z wprowadzonych cegieł. " \
          "W pliku wyjściowym excel program utworzy zakladki odpowiadające wprowadzonym " \
          "nazwom cegieł, a następnie zapisze w nich dane PnVa oraz UNITS dla leków." \
          "\n\n" \
          "\t\t\t\t\t\tWAŻNE !!!" \
          "\n\n" \
          "Wprowadzane nazwy leków oraz cegieł służą do odnajdywania danych PnVa oraz UNITS w pliku excel, " \
          "muszą zatem być identyczne z nazwami znajdującymi się w pliku wejściowym excel. " \
          "Zalecam kopiowanie nazw bezpośrednio z pliku excel do profilu. " \
          "W przypadku niezgodności nazw dane leków nie zostaną odnalezione."

intro_p = " Excel Data Converter - Menadżer profili "


class ProfileMenagerInfo(object):
    def __init__(self):

        info = Tkinter.Tk()
        info.wm_title("Menadżer profili - Informacje")
        info.wm_resizable(width="true", height="true")
        # info.minsize(width=500, height=500)
        # info.maxsize(width=500, height=500)

        label = Tkinter.Label(info, text=Men_info, height=40, justify="left", wraplength=700)
        label.grid(row=0, column=0, padx=10, sticky="ew")
        button = Tkinter.Button(info, text="Zamknij", borderwidth=2, command=info.destroy, padx=20)
        button.grid(row=1, column=0, sticky="nswe")


class ProfileMenager(SQliteEdit):
    def __init__(self):
        SQliteEdit.__init__(self)

        self.profile = Tkinter.Tk()
        self.profile.wm_title("Menadżer profili")
        self.profile.wm_resizable(width="false", height="false")
        sizex, sizey, posx, posy = 1220, 670, 100, 100                         # wymiary okna
        self.profile.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, posx, posy))

        # wymiary sa bez większego znaczenia
        # width=500,height=100    bez znaczenia   bd = 3 - granica!!!
        mainframe = Tkinter.Frame(self.profile, relief="groove", width=500, height=700, bd=4)
        mainframe.grid(row=0, column=0)

        canvas = Tkinter.Canvas(mainframe)      # canvas wewnątrz mainframe
        self.inerframe = Tkinter.Frame(canvas)       # inerframe wewnątrz canvas

        myscrollbar = ttk.Scrollbar(mainframe, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=myscrollbar.set)

        myscrollbar.grid(row=0, column=1, sticky="ns")
        canvas.grid(row=0, column=0)

        def my_function(event):
            # wymiary ruchomej zamki !!!
            # wymiaru musza być mniejsze niż zawartośc frame - czyli ramki której dotyczy scroll
            canvas.configure(scrollregion=canvas.bbox("all"), width=sizex-30, height=sizey-10)

        canvas.create_window((0, 0), window=self.inerframe, anchor='nw')
        self.inerframe.bind("<Configure>", my_function)    # do interframe trzeba wszystko wpakować
        """
        trzeba wprowadzić dużegoframa
        do niego wpakować canvasa
        i do canwasa wewnętrznego frama a do niego pozostałe żeczy
        """
        # profile.grid_columnconfigure(0, weight=1)
        # profile.grid_columnconfigure(1, weight=1)
        # profile.grid_columnconfigure(2, weight=1)

        self.defoult_listA = [["", ""], ["", ""], ["", ""]]
        self.defoult_listB = [["", ""]]

        # self.entrylistA = None
        # self.textlistA = None

        self.entrylistB = None
        self.textlistB = None
        self.e_profile_name = None
        self.com_profile_edit = None
        self.com_profile_del = None
        self.e_cegla_list = None
        self.profile_del = None

        self.nazwa_nowy_profil = "Wprowadź nazwę nowego profilu"
        self.nazwa_cegla_list = "Wprowadź nazwy dla regionów / cegieł"

        self.profiles_name_list = []
        self.profil_PnVA_Units = []
        self.final_profile = []

        self.profile_elem1()
        self.profile_elem2()
        self.profile_elem3()

        self.profile.mainloop()

    def ok(self): pass

    def text_converter(self):
        """ Przetwarza wprowadzony tekst z pół Entry i Text do by można było dodać lub odjąć okna entry """

        self.profil_PnVA_Units = []
        for x in range(len(self.entrylistA)):
            a = []
            nazwagrupy = self.entrylistA[x].get()
            nazwalekow = str(self.textlistA[x].get(1.0, 'end'))[:-1]
            a.append(nazwagrupy)
            a.append(nazwalekow)
            self.profil_PnVA_Units.append(a)

    def dodanie_grupy(self):
        """ Dodaje jeden wiersz pół entry i text po lewej stronie """

        self.text_converter()
        self.defoult_listA.append(["", ""])

        for n in range(len(self.defoult_listA)-1):
            self.defoult_listA[n][0] = self.profil_PnVA_Units[n][0]
            self.defoult_listA[n][1] = self.profil_PnVA_Units[n][1]
        self.profile_elem2()

    def odjecie_grupy(self):
        """ odejmuje jeden wiersz pół entry i text po lewej stronie  """

        self.text_converter()
        self.defoult_listA = self.defoult_listA[:-1]

        if len(self.defoult_listA) == 0:
            self.defoult_listA = [["", ""]]

        for n in range(len(self.defoult_listA)):
            self.defoult_listA[n][0] = self.profil_PnVA_Units[n][0]
            self.defoult_listA[n][1] = self.profil_PnVA_Units[n][1]

        self.profile_elem2()

    def final_text_converter(self):
        """ Przetwarza wprowadzony tekst z pół Entry i Text do formy przystępnej dla SQLite """

        self.final_profile = []
        for x in range(len(self.entrylistA)):
            a = []
            nazwagrupy = self.entrylistA[x].get()
            nazwalekow = str(self.textlistA[x].get(1.0, 'end'))[:-1]
            if len(nazwagrupy) > 1:
                if len(nazwalekow) > 1:
                    a.append("PvNa_UNITS")
                    a.append(nazwagrupy)
                    a.append(nazwalekow)
                    self.final_profile.append(tuple(a))

        a = []
        # nazwagrupy = entrylistB[x].get()
        nazwagrupy = ""
        nazwalekow = str(self.textlistB[0].get(1.0, 'end'))[:-1]
        if len(nazwalekow) > 1:
            a.append("CEGLY")
            a.append(nazwagrupy)
            a.append(nazwalekow)
            self.final_profile.append(tuple(a))
        else:
            a.append("CEGLY")
            a.append(nazwagrupy)
            a.append(" ")
            self.final_profile.append(tuple(a))

        lista_cegiel = ["lista_cegiel", ""]

        list = str(self.e_cegla_list.get())
        lista_cegiel.append(list)
        self.final_profile.append(tuple(lista_cegiel))

        self.final_profile = tuple(self.final_profile)
        print "list", len(list), list


    def create_profile(self):
        """ Zapisuje dane do tabelek w pliku db jako bazy danych SQLite
        WAŻNA NAZWA NIE MOŻE ZAWIERAĆ WEWNĘTRZNYCH SPACJI !!!  """

        daneOK = 0
        try:
            nazwa = str(self.e_profile_name.get())
        except UnicodeEncodeError:
            daneOK += 1
            tkMessageBox.showerror("Błąd...","Nazwa profilu nie może zawierać polskich liter ani spacji.")

        try:
            cegly = str(self.e_cegla_list.get()).strip()
            print "cegly", cegly, "cegly", len(cegly)   ########################################################################3
            print type(cegly)

            if len(cegly)==0:
                print "111"
                daneOK += 1
                tkMessageBox.showerror("Błąd...","Nie wprowadzono prawidłowo nazw dla cegieł.")

        except UnicodeEncodeError:
            print "222"
            daneOK += 1
            tkMessageBox.showerror("Błąd...","Nie wprowadzono prawidłowo nazw dla cegieł.")


        if daneOK == 0:
            self.final_text_converter()
            if len(self.final_profile) <= 2:
                daneOK += 1
                tkMessageBox.showerror("Błąd...","Nie wprowadzono nazw leków w grupach w profilu " + nazwa + ". \n\n"
                                        "Należy wprowadzić nazwy dla grup leków oraz nazwy leków w grupach. \n"
                                        "Przeczytaj również zawartośc okna Instrukcja obsługi Menadżera profili.")
            if self.final_profile[1][2] == " ":
                daneOK += 1
                tkMessageBox.showerror("Błąd...","Nie wprowadzono nazw leków w cegłach w profilu " + nazwa + ". \n\n"
                                        "Należy wprowadzić nazwy dla grup leków oraz nazwy leków w grupach. \n"
                                        "Przeczytaj również zawartośc okna Instrukcja obsługi Menadżera profili.")

        if daneOK == 0:      # zapis do bazy danych
            self.SQsave_profile(nazwa, self.final_profile)    # zapisanie profilu w bazie SQlite
            tkMessageBox.showinfo("Utworzono Profil...",
                                  "Utworzono profil o nazwie: \n " + nazwa + ".\n"
                                  "Zawartość nowo utworzonego profilu " + nazwa + " można od teraz edytować.")
            self.profile_elem1()


    def edit_profile(self):

        profil_to_edit = self.com_profile_edit.get()
        self.defoult_listA = []
        self.defoult_listB = []

        self.SQedit_profile(profil_to_edit)

        self.cursor.execute("SELECT * FROM " + profil_to_edit)
        from_db = self.cursor.fetchall()

        for row in from_db:
            if row[0] == "PvNa_UNITS":
                a = []
                a.append(row[1])
                a.append(row[2:][0])
                self.defoult_listA.append(a)

            if row[0] == "CEGLY":
                a = []
                a.append(row[1])
                a.append(row[2:][0])
                self.defoult_listB.append(a)

            if row[0] == "lista_cegiel":
                self.nazwa_cegla_list = row[2]

        self.nazwa_nowy_profil = profil_to_edit
        self.profile_elem1()
        self.profile_elem2()
        self.profile_elem3()

    def del_profile(self):

        profile_to_del = self.com_profile_del.get()
        self.SQdel_profile(profile_to_del)                  # usunięcie profilu
        tkMessageBox.showinfo("Usunięto Profil...","Usunięto Profil o nazwie:  " + profile_to_del)

        self.profile_elem1()

    ############################################################
    #GUI
    ############################################################

    def profile_elem1(self):
        """    Tworzy poszczególne elementy interface. Cześć 1. """

        self.get_tables_from_db()
        szerokosc = 47
        wysokosc = 5

        # Ramki w ramach GUI
        title = Tkinter.Frame(self.inerframe)
        buttons = Tkinter.Frame(self.inerframe)
        title.grid(row=0, column=0, columnspan=2, sticky="nswe")
        buttons.grid(row=1, column=0, columnspan=2, sticky="we")
        ####################################################################

        l_title = Tkinter.Label(title, text=intro_p)
        b_instruction = Tkinter.Button(title, text="Instrukcja obsługi\nMenadżera profili", bg="orange",
                                       borderwidth=2, command=ProfileMenagerInfo, padx=20)

        l_title.grid(row=0, column=0, columnspan=4, padx=400, sticky="ew")
        b_instruction.grid(row=0, column=5, sticky="nswe")

        ####################################################################
        l_new_profile = Tkinter.Label(buttons, text = "Nazwa nowego Profilu", padx=szerokosc, pady=wysokosc)

        self.e_profile_name = Tkinter.Entry(buttons)
        self.e_profile_name.insert(0, self.nazwa_nowy_profil)

        b_save_profile = Tkinter.Button(buttons, text="ZAPISZ PROFIL", anchor="center", borderwidth=2,
                                        command=self.create_profile, padx = szerokosc, pady=wysokosc)
        ######
        l_profile_to_edit = Tkinter.Label(buttons, text="Nazwa Profilu do edycji", padx=szerokosc, pady=wysokosc)

        self.com_profile_edit = ttk.Combobox(buttons)
        self.com_profile_edit.insert('0', "Wybierz profil do edycji")
        self.com_profile_edit['values'] = self.profiles_name_list

        b_edit_profile = Tkinter.Button(buttons, text="EDYTÓJ PROFIL", padx =szerokosc, pady=wysokosc,
                                        borderwidth=2, command=self.edit_profile)
        ######
        l_profile_to_del = Tkinter.Label(buttons, text="Nazwa Profilu do usunięcia", padx=szerokosc, pady=wysokosc)

        self.com_profile_del = ttk.Combobox(buttons)
        self.com_profile_del.insert('1', "Wybierz profil do usunięcia")
        self.com_profile_del['values'] = self.profiles_name_list

        b_del_profile = Tkinter.Button(buttons, text="USUŃ PROFIL", padx=szerokosc, pady=wysokosc,
                                       borderwidth=2, command=self.del_profile)
        ######
        b_close = Tkinter.Button(buttons, text="ZAMKNIJ EDYTOR PROFILI", padx=szerokosc, borderwidth=2,
                                 command=self.profile.destroy)
        ######
        l_lista_cegiel = Tkinter.Label(buttons, text="Lista regionów / cegieł", padx =szerokosc/2, pady = 10)

        self.e_cegla_list = Tkinter.Entry(buttons)
        self.e_cegla_list.insert(0, self.nazwa_cegla_list)

        ######
        l_new_profile.grid(row=0, column=0, sticky="swen")
        l_profile_to_edit.grid(row=0, column=1, sticky="swen")
        l_profile_to_del.grid(row=0, column=2, sticky="swen")

        self.e_profile_name.grid(row=1, column=0, sticky="we", padx=szerokosc, pady=wysokosc, ipadx=50)
        self.com_profile_edit.grid(row=1, column=1, sticky="we", padx=szerokosc, pady=wysokosc, ipadx=20)
        self.com_profile_del.grid(row=1, column=2, sticky="we", padx=szerokosc, ipadx=20, pady=wysokosc)
        b_close.grid(row=1, column=3, sticky="swen")

        b_save_profile.grid(row=2, column=0, sticky="swen")
        b_edit_profile.grid(row=2, column=1, sticky="swen")
        b_del_profile.grid(row=2, column=2, sticky="swen")

        empty_line = Tkinter.Label(buttons).grid(row=3, column=0, columnspan=3, sticky="snew")

        l_lista_cegiel.grid(row=4, column=0, sticky="e")
        self.e_cegla_list.grid(row=4, column=1, columnspan=3, sticky="we", padx=szerokosc, pady=10)

        empty_line = Tkinter.Label(buttons).grid(row=5, column=0, columnspan=3, sticky="snew")


    def profile_elem2(self):
        """    Tworzy poszczególne elementy interface. Cześć 2. Cześć z polami Entry i text po lewej stronie """

        # Ramka
        pola_nazw1 = ttk.Frame(self.inerframe, width=10)
        pola_nazw1.grid(row=2, column=0, sticky="nswe")

        pola_nazw1_buttons = Tkinter.Frame(pola_nazw1)
        pola_nazw1_buttons.grid(row=0, column=0, columnspan=2, sticky="nswe")
        ########################################

        l_liczba_zakladek = Tkinter.Label(pola_nazw1_buttons, text="Zmiana liczby zakładek")
        b_add_zakladka = Tkinter.Button(pola_nazw1_buttons, text=" + ", command=self.dodanie_grupy, padx=20)
        b_minus_zakladka = Tkinter.Button(pola_nazw1_buttons, text=" - ", command=self.odjecie_grupy,padx=20)

        l_liczba_zakladek.grid(row=0, column=0, sticky="snew", padx=30)
        b_add_zakladka.grid(row=0, column=1, sticky="ns")
        b_minus_zakladka.grid(row=0, column=2, sticky="ns")
        ########################################

        l_zakladka_names = Tkinter.Label(pola_nazw1, text="Nazwy zakładek w excel'u \nw plikach z konkurecją",
                               bg="grey", bd=1, relief="ridge", height=4, padx=1)
        l_nazwy_lekow = Tkinter.Label(pola_nazw1, text="Nazwy leków w zakładce w plikach z konkurencją ",
                               bg="grey", bd=1, relief="ridge", height=4, padx=1)

        l_zakladka_names.grid(row=1, column=0, sticky="nswe")
        l_nazwy_lekow.grid(row=1, column=1, sticky="nswe")

        self.entrylistA = []
        self.textlistA = []

        x = 2          # numer row od którego zaczynają się entry i texty
        for n in range(len(self.defoult_listA)):
            # Entry text
            entry = Tkinter.Entry(pola_nazw1, borderwidth=1, relief="ridge")
            entry.insert(1, self.defoult_listA[n][0])

            # pole tekstowe
            text = Tkinter.Text(pola_nazw1, height=5, width = 70, pady = 2, borderwidth=1, relief="ridge")
            text.insert('1.0', self.defoult_listA[n][1])

            text.grid(row=x+n, column=1, sticky="nswe")
            entry.grid(row=x+n, column=0, sticky="nwe")

            self.entrylistA.append(entry)
            self.textlistA.append(text)


    def profile_elem3(self):
        """    Tworzy poszczególne elementy interface. Cześć 3. Z polami Entry i text po prawej stronie"""

        pola_nazw2 = Tkinter.Frame(self.inerframe)
        pola_nazw2.grid(row=2, column=1, sticky="nswe")

        empty_line = Tkinter.Frame(pola_nazw2, height=26).grid(row=0, column=0)

        l_nazwy_lekow = Tkinter.Label(pola_nazw2, text="Nazwy leków w pliku z podziałem na CEGŁY", bg = "white",
                               borderwidth=1, relief="ridge",height=4)
        l_nazwy_lekow.grid(row=1, column=0, sticky="nswe")

        self.entrylistB = []
        self.textlistB = []

        t_leki = Tkinter.Text(pola_nazw2, height=16, width=58, borderwidth=1, relief="ridge")
        t_leki.grid(row=2, column=0, sticky="nswe")
        t_leki.insert('1.0', self.defoult_listB[0][1])

        self.textlistB.append(t_leki)

if __name__ == "__main__":
    ProfileMenager()
