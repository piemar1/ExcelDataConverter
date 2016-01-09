# ! /usr/bin/env python
# -*- coding: utf-8 -*-

import sqlite3 as lite
from edc_sqlite import SQliteEdit

import Tkinter
import tkMessageBox
import ttk
import sys
import tkFileDialog
import os

__author__ = 'Marcin Pieczyński'

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
          "Należy wprowadzić nazwy dla poszczególnych zakładek/gup leków - " \
           "nazwy grup staną się w pliku wyjściowym excel " \
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


class ProfileManagerInfo(object):
    """
    The window with information about usage the Profile Manager.
    """
    def __init__(self):

        info = Tkinter.Tk()
        info.wm_title("Menadżer profili - Informacje")
        info.wm_resizable(width="true", height="true")

        label = Tkinter.Label(info, text=Men_info, height=40, justify="left", wraplength=800)
        label.grid(row=0, column=0, padx=10, sticky="ew")
        button = Tkinter.Button(info, text="Zamknij", borderwidth=2, command=info.destroy, padx=20)
        button.grid(row=1, column=0, sticky="nswe")


class ProfileManager(SQliteEdit):
    """
    The Class containing methods for creating and using ProfileManager.
    """
    def __init__(self):
        SQliteEdit.__init__(self)

        self.profile = Tkinter.Tk()
        self.profile.wm_title("Menadżer profili")
        self.profile.wm_resizable(width="false", height="false")
        sizex, sizey, posx, posy = 1220, 670, 100, 100                         # Window size
        self.profile.wm_geometry("%dx%d+%d+%d" % (sizex, sizey, posx, posy))

        mainframe = Tkinter.Frame(self.profile, relief="groove", width=500, height=700, bd=4)
        mainframe.grid(row=0, column=0)

        canvas = Tkinter.Canvas(mainframe)          # canvas inside mainframe
        self.inerframe = Tkinter.Frame(canvas)      # inerframe inside canvas

        myscrollbar = ttk.Scrollbar(mainframe, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=myscrollbar.set)

        myscrollbar.grid(row=0, column=1, sticky="ns")
        canvas.grid(row=0, column=0)

        def my_function(event):
            # size of movable frame
            # dimensions must be a bit smaller than frame
            canvas.configure(scrollregion=canvas.bbox("all"), width=sizex-30, height=sizey-10)

        canvas.create_window((0, 0), window=self.inerframe, anchor='nw')
        self.inerframe.bind("<Configure>", my_function)    # elements of GUI inside of interframe

        self.defaults_listA = [["", ""], ["", ""], ["", ""]]
        self.defaults_listB = [["", ""]]

        self.new_profile_name = "Wprowadź nazwę nowego profilu"
        self.list_ceglas_list = "Wprowadź nazwy dla regionów / cegieł"
        self.profiles_name_list = []
        self.profile_PnVA_Units = []
        self.final_profile = []

        self.entry_listB = None
        self.text_listB = None
        self.e_profile_name = None
        self.com_profile_edit = None
        self.com_profile_del = None
        self.e_cegla_list = None
        self.profile_del = None
        self.entrylistA = None
        self.textlistA = None
        self.entry_listB = None
        self.text_listB = None

        self.profile_elem1()
        self.profile_elem2()
        self.profile_elem3()

        self.profile.mainloop()

    def text_converter(self):
        """
        Converting text from Entry and Text fields for adong or removing of windows.
        """
        self.profile_PnVA_Units = []

        for x in range(len(self.entrylistA)):
            nazwagrupy = self.entrylistA[x].get()
            nazwalekow = str(self.textlistA[x].get(1.0, 'end'))[:-1]

            self.profile_PnVA_Units.append([nazwagrupy, nazwalekow])

    def dodanie_grupy(self):
        """
        The methods for adding one row for Entry and Text field.
        """
        self.text_converter()
        self.defaults_listA.append(["", ""])

        for n in range(len(self.defaults_listA)-1):
            self.defaults_listA[n][0] = self.profile_PnVA_Units[n][0]
            self.defaults_listA[n][1] = self.profile_PnVA_Units[n][1]
        self.profile_elem2()

    def odjecie_grupy(self):
        """
        The methods for removing one row for Entry and Text field.
        """
        self.text_converter()
        self.defaults_listA = self.defaults_listA[:-1]

        if len(self.defaults_listA) == 0:
            self.defaults_listA = [["", ""]]

        for n in range(len(self.defaults_listA)):
            self.defaults_listA[n][0] = self.profile_PnVA_Units[n][0]
            self.defaults_listA[n][1] = self.profile_PnVA_Units[n][1]

        self.profile_elem2()

    def final_text_converter(self):
        """
        Converting text from Entry and Text fields for saving profile in SQlite
        """
        self.final_profile = []
        for x in range(len(self.entrylistA)):
            nazwagrupy = self.entrylistA[x].get()
            nazwalekow = str(self.textlistA[x].get(1.0, 'end'))[:-1]

            if len(nazwagrupy) > 1 and len(nazwalekow) > 1:
                a = ["PvNa_UNITS", nazwagrupy, nazwalekow]
                self.final_profile.append(tuple(a))

        nazwagrupy = ""
        nazwalekow = str(self.text_listB[0].get(1.0, 'end'))[:-1]
        if len(nazwalekow) > 1:
            a = ["CEGLY", nazwagrupy, nazwalekow]
        else:
            a = ["CEGLY", nazwagrupy, " "]
        self.final_profile.append(tuple(a))

        lista_cegiel = ["lista_cegiel", "", str(self.e_cegla_list.get())]

        self.final_profile.append(tuple(lista_cegiel))
        self.final_profile = tuple(self.final_profile)

    @staticmethod
    def error_wrong_profile_name():
        """
        The window with message informing that the profile name is incorrect.
        """
        tkMessageBox.showerror("Błąd...", "Nazwa profilu nie może zawierać polskich liter ani spacji.")

    @staticmethod
    def error_wrong_ceglas_names():
        """
        The window with message informing that the ceglas names are incorrect.
        """
        tkMessageBox.showerror("Błąd...", "Nie wprowadzono prawidłowo nazw dla cegieł.")

    @staticmethod
    def error_no_drugs_names(name):
        """
        The window with message informing that drugs names have not been introduced.
        """
        tkMessageBox.showerror("Błąd...", "Nie wprowadzono nazw leków w grupach w profilu " + name + ". \n\n"
                               "Należy wprowadzić nazwy dla grup leków oraz nazwy leków w grupach. \n"
                               "Przeczytaj również zawartośc okna Instrukcja obsługi Menadżera profili.")

    def create_profile(self):
        """
        The method for saving profile into SQLite database file.
        """
        try:
            name = str(self.e_profile_name.get())
        except UnicodeEncodeError:
            return self.error_wrong_profile_name()

        try:
            if len(str(self.e_cegla_list.get()).strip()) == 0:
                return self.error_wrong_ceglas_names()
        except UnicodeEncodeError:
                return self.error_wrong_ceglas_names()

        self.final_text_converter()

        if len(self.final_profile) <= 2 or self.final_profile[1][2] == " ":
            return self.error_no_drugs_names(name)

        self.sqsave_profile(name, self.final_profile)    # saving profile to SQlite database file

        tkMessageBox.showinfo("Utworzono Profil...",
                              "Utworzono profil o nazwie: \n " + name + ".\n"
                              "Zawartość nowo utworzonego profilu " + name + " można od teraz edytować.")
        self.profile_elem1()

    def edit_profile(self):
        """
        The method for reading profile from SQLite database file for edition.
        """
        profile_to_edit = self.com_profile_edit.get()
        self.defaults_listA, self.defaults_listB = [], []

        self.cursor.execute("SELECT * FROM " + profile_to_edit)
        from_db = self.cursor.fetchall()

        for row in from_db:
            if row[0] == "PvNa_UNITS":
                a = [row[1], row[2:][0]]
                self.defaults_listA.append(a)

            if row[0] == "CEGLY":
                a = [row[1], row[2:][0]]
                self.defaults_listB.append(a)

            if row[0] == "lista_cegiel":
                self.list_ceglas_list = row[2]

        self.new_profile_name = profile_to_edit
        self.profile_elem1()
        self.profile_elem2()
        self.profile_elem3()

    def del_profile(self):
        """
        The method for deleting profile from SQLite database.
        """
        profile_to_del = self.com_profile_del.get()
        self.sqdel_profile(profile_to_del)
        tkMessageBox.showinfo("Usunięto Profil...", "Usunięto Profil o nazwie:  " + profile_to_del)

        self.profile_elem1()

    #####
    # GUI

    def profile_elem1(self):
        """
        The methods creating first (upper) part of GUI elements.
        """
        self.get_tables_from_db()
        szerokosc, wysokosc = 47, 5

        # Frames in GUI
        title = Tkinter.Frame(self.inerframe)
        buttons = Tkinter.Frame(self.inerframe)
        title.grid(row=0, column=0, columnspan=2, sticky="nswe")
        buttons.grid(row=1, column=0, columnspan=2, sticky="we")
        ####################################################################

        l_title = Tkinter.Label(title, text=intro_p)
        b_instruction = Tkinter.Button(title, text="Instrukcja obsługi\nMenadżera profili", bg="orange",
                                       borderwidth=2, command=ProfileManagerInfo, padx=20)

        l_title.grid(row=0, column=0, columnspan=4, padx=400, sticky="ew")
        b_instruction.grid(row=0, column=5, sticky="nswe")

        ####################################################################
        l_new_profile = Tkinter.Label(buttons, text="Nazwa nowego Profilu", padx=szerokosc, pady=wysokosc)

        self.e_profile_name = Tkinter.Entry(buttons)
        self.e_profile_name.insert(0, self.new_profile_name)

        b_save_profile = Tkinter.Button(buttons, text="ZAPISZ PROFIL", anchor="center", borderwidth=2,
                                        command=self.create_profile, padx=szerokosc, pady=wysokosc)
        ######
        l_profile_to_edit = Tkinter.Label(buttons, text="Nazwa Profilu do edycji", padx=szerokosc, pady=wysokosc)

        self.com_profile_edit = ttk.Combobox(buttons)
        self.com_profile_edit.insert('0', "Wybierz profil do edycji")
        self.com_profile_edit['values'] = self.profiles_name_list

        b_edit_profile = Tkinter.Button(buttons, text="EDYTÓJ PROFIL", padx=szerokosc, pady=wysokosc,
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
        l_lista_cegiel = Tkinter.Label(buttons, text="Lista regionów / cegieł", padx=szerokosc/2, pady=10)

        self.e_cegla_list = Tkinter.Entry(buttons)
        self.e_cegla_list.insert(0, self.list_ceglas_list)

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
        """
        The methods creating second (lower left) part of GUI elements with Entry and Text fields.
        """
        # Frame
        pola_nazw1 = ttk.Frame(self.inerframe, width=10)
        pola_nazw1.grid(row=2, column=0, sticky="nswe")

        pola_nazw1_buttons = Tkinter.Frame(pola_nazw1)
        pola_nazw1_buttons.grid(row=0, column=0, columnspan=2, sticky="nswe")
        #####

        l_liczba_zakladek = Tkinter.Label(pola_nazw1_buttons, text="Zmiana liczby zakładek")
        b_add_zakladka = Tkinter.Button(pola_nazw1_buttons, text=" + ", command=self.dodanie_grupy, padx=20)
        b_minus_zakladka = Tkinter.Button(pola_nazw1_buttons, text=" - ", command=self.odjecie_grupy, padx=20)

        l_liczba_zakladek.grid(row=0, column=0, sticky="snew", padx=30)
        b_add_zakladka.grid(row=0, column=1, sticky="ns")
        b_minus_zakladka.grid(row=0, column=2, sticky="ns")
        #####

        l_zakladka_names = Tkinter.Label(pola_nazw1, text="Nazwy zakładek w excel'u \nw plikach z konkurecją",
                                         bg="grey", bd=1, relief="ridge", height=4, padx=1)
        l_nazwy_lekow = Tkinter.Label(pola_nazw1, text="Nazwy leków w zakładce w plikach z konkurencją ",
                                      bg="grey", bd=1, relief="ridge", height=4, padx=1)

        l_zakladka_names.grid(row=1, column=0, sticky="nswe")
        l_nazwy_lekow.grid(row=1, column=1, sticky="nswe")

        self.entrylistA, self.textlistA = [], []

        x = 2          # 2 is a row number for Entry and Text fields
        for n in range(len(self.defaults_listA)):
            # Entry
            entry = Tkinter.Entry(pola_nazw1, borderwidth=1, relief="ridge")
            entry.insert(1, self.defaults_listA[n][0])

            # Text
            text = Tkinter.Text(pola_nazw1, height=5, width=70, pady=2, borderwidth=1, relief="ridge")
            text.insert('1.0', self.defaults_listA[n][1])

            text.grid(row=x+n, column=1, sticky="nswe")
            entry.grid(row=x+n, column=0, sticky="nwe")

            self.entrylistA.append(entry)
            self.textlistA.append(text)

    def profile_elem3(self):
        """
        The methods creating third (lower right) part of GUI elements with Entry and Text fields.
        """
        pola_nazw2 = Tkinter.Frame(self.inerframe)
        pola_nazw2.grid(row=2, column=1, sticky="nswe")

        empty_line = Tkinter.Frame(pola_nazw2, height=26).grid(row=0, column=0)

        l_nazwy_lekow = Tkinter.Label(pola_nazw2, text="Nazwy leków w pliku z podziałem na CEGŁY", bg="white",
                                      borderwidth=1, relief="ridge", height=4)
        l_nazwy_lekow.grid(row=1, column=0, sticky="nswe")

        self.entry_listB, self.text_listB = [], []

        t_leki = Tkinter.Text(pola_nazw2, height=16, width=58, borderwidth=1, relief="ridge")
        t_leki.grid(row=2, column=0, sticky="nswe")
        t_leki.insert('1.0', self.defaults_listA[0][1])

        self.text_listB.append(t_leki)

if __name__ == "__main__":
    ProfileManager()
