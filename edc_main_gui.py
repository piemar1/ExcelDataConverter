# ! /usr/bin/env python
# -*- coding: utf-8 -*-
__author__ = 'Marcin Pieczyński'

import sqlite3 as lite
import Tkinter
import ttk
import tkMessageBox
import os
import tkFileDialog

from edc_sqlite import SQliteEdit
from edc_init import LicenceCheck
from edc_read_input import Read_input

import EDC_write_output
import edc_profile_manager

# import read_input
# import write_output

# print globals()
# print "#######################################################3"
# print dir(EDC_init)
# for k, v in globals().iteritems():
#     print k, v

# style_bold8 = easyxf('font: name Tahoma, bold on, height 160;')
# style_no_bold = easyxf('font: name Tahoma, height 160;')


top_info = """\t\t\t\t\tEXCEL Data Converter\n\n"\
           "Program służy do filtrowania danych z pliku wejściowego excel i ich zapisu w plikach wyjściowych excel.\n\n" \
           "W celu filtrowania danych należy dokonać wyboru pliku wejściowego z danymi excel oraz trzech plików " \
           "wyjściowych, w których dane zostaną zapisane. Należy również wybrać profil filtrowania - czyli sposób " \
           "w jaki program będzie filtrować dane. \n\n" \
           "W celu utworzenia lub edycji profilu filtrowania danych należy skożystać z Menadżera Profili oraz " \
           "zawartej w nim instrukcji. " \
           "Po dokonaniu zmian w profilach filtrowania należy odświerzyć listę profili przed dalszym działaniem. " \
           "Wybrany profil można następnie wybrać z listy. Plik wejściowy powinnien stanowić plik xlsx zawierający " \
           "dane, które zamierzamy przefiltrować a następnie zapisać w 3 plikach wyjściowych excel. \n\n" \
           "Plik PnVa będzie zawierał dane PnVa dla wybranych leków w wybranych cegłach, " \
           "natomiast plik UNITS będzie zawierał dane UNITS, również dla wybranych leków i cegieł. " \
           "Plik CEGŁY będzie natomiast zawierał dane PnVa oraz UNITS dla wybranych leków pogrupowane " \
           "w wybranych cegłach.\n\n" \
           "Podczas wpisywania nazw plików wyjściowych " \
           "należy jedynie wpisać nazwy plików bez rozwinięć plików w postaci koncówki '.xls'. " \
           "Rozwinięcia będą dodawane automatycznie. Program konwertuje dane w czasie od kilku " \
           "do kilkudziesięciu sekund - w zależności od liczby leków wprowadzonych w profilu. \n\n" \
           "\t\t\t\tUżytkowanie programu w przyszłości. \n\n" \
           "Obecna wersja programu jest jedynie jednomiesięczną wersją testową. " \
           "W celu korzystanie z programu w przyszłości możliwe będzie wykupienie za niewielką opłatą " \
           "miesięcznej licencji na użytkowanie programu. Licencja ważna będzie na jeden komputer. " \
           "Możliwe będzie również przenoszenie utwożonych profili filtrowania danych na kolejne wersje programu. " \
           "W przypadku zainteresowania dalszym użytkowaniem zamierzam w przyszłości rozwijać i aktualizować program."\
           "\n\n" \
           "Przykładowe zmiany jakie można wprowadzić w następnych wersjach programu:\n" \
           "- inny sposób filtrowania danych\n" \
           "- inny sposób ułożenia i formatowania danych w plikach wyjściowych\n" \
           "- dodanie opcji tworzenia wykresów dla wybranych danych\n" \
           "- poprawienie oprawy graficznej\n" \
           "- inne zmiany zaproponowane przez użytkowników - " \
           "Excel ma wiele możliwośc by przetważać i przedstawiać informacje według potrzeb, " \
           "czemu tego nie zautomatyzować? \n\n"\
           "W przypadku pytań lub problemów chętnię pomogę. \n\n"\
           "\t\t\t\t\t      Marcin Pieczyński \n\t\t\t\t\t marcin-pieczynski@wp.pl"""

class InformacjaTop():
    def __init__(self):
        info = Tkinter.Tk()
        info.wm_title("Excel Data Converter - Informacje")
        info.wm_resizable(width="true", height="true")
        # info.minsize(width=500, height=500)
        # info.maxsize(width=500, height=500)

        label = Tkinter.Label(info, text=top_info, height=42, justify="left", wraplength=600)
        label.grid(row=0, column=0, sticky="ew")

        Button = Tkinter.Button(info, text="Zamknij", borderwidth=2, command=info.destroy, padx=20)
        Button.grid(row=1, column=0, sticky="nswe")


class MainGui(LicenceCheck, SQliteEdit):
    def __init__(self):

        LicenceCheck.__init__(self)
        SQliteEdit.__init__(self)

        self.top = Tkinter.Tk()
        self.top.wm_title("Excel Data Converter")
        self.top.wm_resizable(width="False", height="False")
        self.top.minsize(width=1000, height=400)
        self.top.maxsize(width=1000, height=400)
        self.top.resizable(width=True, height=False)

        self.intro = "\nExcel Data Converter\n"

        self.imputnumber, self.PvNanumber, self.UNITSnumber, self.Ceglynumber = 0, 0, 0, 0
        self.filepath_input = Tkinter.StringVar()
        self.filepath_PnVa = Tkinter.StringVar()
        self.filepath_UNITS = Tkinter.StringVar()
        self.filepath_CEGLY = Tkinter.StringVar()

        self.licence_days = "Do końca licencji pozostało " + str(self.delta) + " dni."

        self.progres_step = 0

        self.interface_elem()
        self.top.mainloop()

    def start_progres_bar(self):
        progres["value"] = self.progres_step
        # print "step", step

    def ad_step_to_progres_bar(self,n):
        self.progres_step += n
        progres["value"] = self.progres_step
        progres.update_idletasks()
        # print "step", step


    def input_file(self):
        """ okno wprowadzające plik z danych wejściowymi"""
        try:
            f = tkFileDialog.askopenfilename(parent=self.top, initialdir="/home/marcin/pulpit/Py/",
                                             title="Wybór pliku excel z danymi", filetypes=[("Excel file", ".xlsx")])
            self.filepath_input.set(os.path.realpath(f))
            self.excel_input_file = os.path.realpath(f)
            self.imputnumber += 1
        except ValueError: pass        #!!!!!!!!!!!! Dodać Error związany z plikiem !!!!!

    def save_filePnVa(self):
        """ Okno wprowadzające dane dla pliku z danych wyjściowymi PnVa"""
        try:
            save = tkFileDialog.asksaveasfilename(parent=self.top, initialdir="/home/marcin/pulpit/",
                                         title="Wybór pliku do zapisu danych PnVa", filetypes=[("Excel file", ".xlsx")])
            self.filepath_PnVa.set((os.path.realpath(save))+".xlsx")
            self.PnVa_file = os.path.realpath(save)+".xlsx"
            self.PvNanumber += 1
        except ValueError: pass

    def save_fileUNITS(self):
        """ Okno wprowadzające dane dla pliku z danych wyjściowymi UNITS"""

        try:
            save = tkFileDialog.asksaveasfilename(parent=self.top, initialdir="/home/marcin/pulpit/",
                                         title="Wybór pliku do zapisu danych UNITS",
                                         filetypes = [("Excel file", ".xlsx")])
            self.filepath_UNITS.set((os.path.realpath(save))+".xlsx")
            self.UNITS_file = os.path.realpath(save)+".xlsx"
            self.UNITSnumber += 1
        except ValueError: pass

    def save_fileCEGLY(self):
        """ Okno wprowadzające dane dla pliku z danych wyjściowymi UNITS"""
        try:
            save = tkFileDialog.asksaveasfilename(parent=self.top, initialdir="/home/marcin/pulpit/",
                                         title="Wybór pliku do zapisu danych z Cegieł",
                                         filetypes = [("Excel file", ".xlsx")])
            # fc = open(save,"w")
            self.filepath_CEGLY.set((os.path.realpath(save))+".xlsx")
            self.CEGLY_file = os.path.realpath(save)+".xlsx"
            self.Ceglynumber += 1
        except ValueError: pass

    def open_profile_menager(self):
        """ importuje i uruchamia menadżer profili"""
        reload(edc_profile_manager)
        edc_profile_manager.ProfileMenager()

    def KRKA_data_convert(self):
        """ Funkcja wywołująca całościową konwersję danych uruchamiająca wiele innych funkcji"""

        profil = self.com_choosen_profile.get()

        if profil == "Wybierz profil":
            tkMessageBox.showerror("Błąd, Nie wybrano profilu.",
                                   "Należy dokonać wyboru profilu do filtrowania danych, "
                                   "szczegóły poszczególnych profili można znaleść w Menadżeże Profili")
        print profil

        if self.imputnumber == 0 or self.PvNanumber == 0 or self.UNITSnumber == 0 or self.Ceglynumber == 0:
            ##
            #    NAleży usunąć te cyferki i zostawić nazwy !!!!!!!!!!!
            #
                    tkMessageBox.showerror(" Błąd, Nie wybrano nazw plików. ",
                                           "Przed konwersją danych należy wybrać plik wejściowy zawierające dane Excel "
                                           "oraz podać nazwy dla 3 plików wyjściowych w których program zapisze dane.")

        if self.imputnumber and self.PvNanumber and self.UNITSnumber and self.Ceglynumber \
                and profil != "Wybierz profil":

            tkMessageBox.showinfo("Yes... do roboty !!!",
                                  "Profil filtrowania danych i pliki wybrane ... przystępujemy do konwersji danych...")

            # FUNKCJA wykonująca właściwą konwersję danych !!!

            # wczytanie pliku wejściowego oraz zawartych w nich danych
            excel_input_file = Read_input(self.excel_input_file)


            # self.start_progres_bar()                                            #    ;ad_step(5)
            # get_data()                                             ;ad_step(15)
            # start_input(excel_input_file)                          ;ad_step(10)
            # # date_finding()                                         ;ad_step(10)
            # KRKA_excel_data_converter(PnVa_file, "Pn Va")          ;ad_step(20)
            # KRKA_excel_data_converter(UNITS_file, "UNITS REPORT")  ;ad_step(20)
            # KRKA_data_converter_cegly(CEGLY_file)                  ;ad_step(20)

            tkMessageBox.showinfo("Yes...", "Dokonano konwersji danych. \n Życzę miłego dnia.")

    # Główne GUI
    #########################################3

    def interface_elem(self):
        """ Tworzenie poszczególnych elementów GUI"""

        self.get_tables_from_db()

        # tekst intro
        l_intro = Tkinter.Label(self.top, text=self.intro, relief="ridge", pady=2, padx=400)

        b_instruction = Tkinter.Button(self.top, text="Instrukcja obsługi programu", borderwidth=2, bg="orange",
                                command= InformacjaTop, pady=5, padx=20)
        # ==========================================================
        buttons = Tkinter.Frame(self.top).grid(row=1, column=0, columnspan=9, sticky="nswe")

        b_profile_menager = Tkinter.Button(buttons, text= "Otwórz Menadżer Profili", padx=30, pady=10,
                                        command=self.open_profile_menager)

        b_profile_reload = Tkinter.Button(buttons, text="Odśwież listę profili",
                                          padx=30, pady=10, command=self.interface_elem)

        self.com_choosen_profile = ttk.Combobox(buttons)
        self.com_choosen_profile.insert('0', "Wybierz profil")
        self.com_choosen_profile['values'] = self.profiles_name_list

        # ==========================================================
        # szukanie pliku wejsciowego
        b_input_file = Tkinter.Button(self.top, text="Plik wejsciowy", command=self.input_file, padx=60)
        l_input_file = Tkinter.Label(self.top, width=7, textvariable=self.filepath_input)

        # plik wyjściowy PvNa
        b_output_PnVa = Tkinter.Button(self.top, text="Plik wyjściowy   PnVa", command=self.save_filePnVa)
        l_output_PnVa = Tkinter.Label(self.top, width=7, textvariable=self.filepath_PnVa)

        # plik wyjściowy UNITS
        b_output_UNITS = Tkinter.Button(self.top, text="Plik wyjściowy   UNITS", command=self.save_fileUNITS)
        l_output_UNITS = Tkinter.Label(self.top, width=7, textvariable=self.filepath_UNITS)

        # plik wyjściowy CEGŁY
        b_output_cegly = Tkinter.Button(self.top, text="Plik wyjściowy   CEGŁY", command=self.save_fileCEGLY)
        l_output_cegly = Tkinter.Label(self.top, width=7, textvariable=self.filepath_CEGLY)

        # przycisk konwertowania danych !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        b_convert = Tkinter.Button(self.top, text="Konwertuj dane !!!", command=self.KRKA_data_convert)
        progres_bar = ttk.Progressbar(self.top, orient= "horizontal", mode='determinate', maximum=100, length=250)

        # Liczba dnia do końca licencji
        l_licence_info = Tkinter.Label(self.top, width=1, text=self.licence_days)

        # zamkniecie programu
        b_close = Tkinter.Button(self.top, text="Zamknij program", command=self.top.destroy,
                               borderwidth=2, relief="ridge", pady=4, padx=30)

        # Połorzenie poszczególnych elementów GUI
        l_intro.grid(row=0, column=0, columnspan=5, sticky="nswe")
        b_instruction.grid(row=0, column=5, sticky="nswe")

        b_profile_menager.grid(row=1, column=1, sticky="we")
        b_profile_reload.grid(row=1, column=3, sticky="we")
        self.com_choosen_profile.grid(row=1, column=4, sticky="we", padx=50, pady = 10, ipadx = 20)
        l_empty_line = Tkinter.Label(buttons).grid(row=1, column=2)
        l_empty_line2 = Tkinter.Label(buttons).grid(row=2, column=0, columnspan=5)

        b_input_file.grid(row=2, column=0, sticky="nswe")
        l_input_file.grid(row=2, column=1, columnspan=10, sticky=("we"))
        l_empty_line3 = Tkinter.Label(self.top).grid(row=3, column=0, columnspan=5)

        b_output_PnVa.grid(row=4, column=0, sticky="nswe")
        l_output_PnVa.grid(row=4, column=1, columnspan=10, sticky=("we"))
        b_output_UNITS.grid(row=5, column=0, sticky="nswe")
        l_output_UNITS.grid(row=5, column=1, columnspan=10, sticky=("we"))
        b_output_cegly.grid(row=6, column=0, sticky="nswe")
        l_output_cegly.grid(row=6, column=1, columnspan=10, sticky=("we"))

        l_empty_line4 = Tkinter.Label(self.top).grid(row=7, column=0, columnspan=5)
        b_convert.grid(row=8, column=3, sticky="nswe")

        l_empty_line5 = Tkinter.Label(self.top).grid(row=9, column=0, columnspan=5)
        progres_bar.grid(row=10, column=3, columnspan=1)

        empty_line6 = Tkinter.Label(self.top).grid(row=11, column=0, columnspan=5)

        l_licence_info.grid(row=12, column=0, columnspan=2, sticky=("we"))
        b_close.grid(row=12, column=5, sticky="nswe")

        for x in range(9):
            self.top.grid_columnconfigure(x,weight=1)

        return progres_bar, self.com_choosen_profile



if __name__ == '__main__':

    lic = edc_init.LicenceCheck()

    MainGui()
    Tkinter.mainloop()






