# ! /usr/bin/env python
# -*- coding: utf-8 -*-

import sqlite3 as lite
import Tkinter
import ttk
import tkMessageBox
import os
import tkFileDialog

from edc_sqlite import SQliteEdit
from edc_init import LicenceCheck
from edc_read_input import ReadInput

from edc_write_output import Write_output_PnVa_Units_paternA, Write_output_CEGLY_paternA
import edc_profile_manager

__author__ = 'Marcin Pieczyński'

# style_bold8 = easyxf('font: name Tahoma, bold on, height 160;')
# style_no_bold = easyxf('font: name Tahoma, height 160;')

top_info = "\t\t\t\t\tEXCEL Data Converter\n\n"\
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
           "\t\t\t\t\t      Marcin Pieczyński \n\t\t\t\t\t marcin-pieczynski@wp.pl"


class InformacjaTop(object):
    """
    Window with information about usage the program.
    """
    def __init__(self):
        info = Tkinter.Tk()
        info.wm_title("Excel Data Converter - Informacje")
        info.wm_resizable(width="true", height="true")
        # info.minsize(width=500, height=500)
        # info.maxsize(width=500, height=500)

        label = Tkinter.Label(info, text=top_info, height=42, justify="left", wraplength=600)
        label.grid(row=0, column=0, sticky="ew")

        button = Tkinter.Button(info, text="Zamknij", borderwidth=2, command=info.destroy, padx=20)
        button.grid(row=1, column=0, sticky="nswe")


class MainGui(LicenceCheck, SQliteEdit):
    """Class containing main gui of the program and all methods connected with the gui."""

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
        """Start of the progres bar."""
        self.progres["value"] = self.progres_step

    def ad_step_to_progres_bar(self, n):
        """Update of the progres bar."""
        self.progres_step += n
        self.progres["value"] = self.progres_step
        self.progres.update_idletasks()

    def input_file(self):
        """ Method for introduction an input file with data for filtration."""
        try:
            f = tkFileDialog.askopenfilename(parent=self.top, initialdir="/home/marcin/pulpit/Py/",
                                             title="Wybór pliku excel z danymi", filetypes=[("Excel file", ".xlsx")])
            self.filepath_input.set(os.path.realpath(f))
            self.excel_input_file = os.path.realpath(f)
            self.imputnumber += 1
        except ValueError:
            tkMessageBox.showerror("Error", " Wystąpił problem z załadowaniem pliku excel z danymi.")

    def save_filePnVa(self):
        """ Method for introduction an output file with PnVa data. """
        try:
            save = tkFileDialog.asksaveasfilename(parent=self.top, initialdir="/home/marcin/pulpit/",
                                                  title="Wybór pliku do zapisu danych PnVa",
                                                  filetypes=[("Excel file", ".xlsx")])
            self.filepath_PnVa.set((os.path.realpath(save)))
            self.PnVa_file = os.path.realpath(save)
            self.PvNanumber += 1
        except ValueError:
            tkMessageBox.showerror("Error", " Wystąpił problem z plikiem do zapisu danych PnVa.")

    def save_fileUNITS(self):
        """ Method for introduction an output file with UNITS data. """
        try:
            save = tkFileDialog.asksaveasfilename(parent=self.top, initialdir="/home/marcin/pulpit/",
                                                  title="Wybór pliku do zapisu danych UNITS",
                                                  filetypes = [("Excel file", ".xlsx")])
            self.filepath_UNITS.set((os.path.realpath(save)))
            self.UNITS_file = os.path.realpath(save)
            self.UNITSnumber += 1
        except ValueError:
            tkMessageBox.showerror("Error", " Wystąpił problem z plikiem do zapisu danych UNITS.")

    def save_fileCEGLY(self):
        """ Method for introduction an output file with area data. """
        try:
            save = tkFileDialog.asksaveasfilename(parent=self.top, initialdir="/home/marcin/pulpit/",
                                                  title="Wybór pliku do zapisu danych z Cegieł",
                                                  filetypes=[("Excel file", ".xlsx")])
            # fc = open(save,"w")
            self.filepath_CEGLY.set((os.path.realpath(save)))
            self.CEGLY_file = os.path.realpath(save)
            self.Ceglynumber += 1
        except ValueError:
            tkMessageBox.showerror("Error", " Wystąpił problem z plikiem do zapisu danych z Cegieł.")

    def open_profile_menager(self):
        """ Methods lounching profile manager."""
        reload(edc_profile_manager)
        edc_profile_manager.ProfileMenager()

    def error_no_profile(self):
        tkMessageBox.showerror("Błąd, Nie wybrano profilu.",
                               "Należy dokonać wyboru profilu do filtrowania danych, "
                               "szczegóły poszczególnych profili można znaleść w Menadżeże Profili")

    def error_no_files(self):
        tkMessageBox.showerror(" Błąd, Nie wybrano nazw plików. ",
                       "Przed konwersją danych należy wybrać plik wejściowy zawierające dane Excel "
                       "oraz podać nazwy dla 3 plików wyjściowych w których program zapisze dane.")

    def work_finished(self):
        tkMessageBox.showinfo("Yes...", "Dokonano konwersji danych. \n Życzę miłego dnia.")

    def KRKA_data_convert(self):
        """ Important !  Method for carrying out reading, filtration and saving of excel data. """

        profil = self.com_choosen_profile.get()
        # print profil

        if profil == "Wybierz profil":
            return self.error_no_profile()

        if self.imputnumber == 0 or self.PvNanumber == 0 or self.UNITSnumber == 0 or self.Ceglynumber == 0:
            return self.error_no_files()

        if self.imputnumber and self.PvNanumber and self.UNITSnumber and self.Ceglynumber \
                and profil != "Wybierz profil":
            tkMessageBox.showinfo("Yes... do roboty !!!",
                                  "Profil filtrowania danych i pliki wybrane ... przystępujemy do konwersji danych...")

    ###############################################################
    # Functions for for carrying out reading, filtration and saving of excel data.

    # Reading of input data for excel file.

            self.start_progres_bar()

            excel_input_file = ReadInput()                          ;self.ad_step_to_progres_bar(5)
            excel_input_file.open_input_file(self.excel_input_file) ;self.ad_step_to_progres_bar(5)
            excel_input_file.get_date()                             ;self.ad_step_to_progres_bar(5)
            excel_input_file.get_PnVa_Units_column()                ;self.ad_step_to_progres_bar(5)
            excel_input_file.get_cegla_positions()                  ;self.ad_step_to_progres_bar(5)
            excel_input_file.get_Cegly_data(self.com_choosen_profile.get())
            self.ad_step_to_progres_bar(5)
            excel_input_file.get_PnVa_UNITS_data(self.com_choosen_profile.get())
            self.ad_step_to_progres_bar(5)
            print "Odczytanie danych z pliku input OK"

            # print "excel_input_file.PnVa_data", excel_input_file.PnVa_data
            # print 50 * "%%%"
            # print "excel_input_file.UNITS_data", excel_input_file.UNITS_data
            print 50 * "%%%"
            print "excel_input_file.CEGLY_data", excel_input_file.CEGLY_data
            print 50 * "%%%"
            # print "excel_input_file.output_zakladki", excel_input_file.output_zakladki
            # print "excel_input_file.output_lista_cegiel", excel_input_file.output_lista_cegiel
            # print "excel_input_file.output_leki", excel_input_file.output_leki


            print "START zapis danych PnVa do pliku output"
            output1 = Write_output_PnVa_Units_paternA(excel_input_file.output_zakladki,
                                                      excel_input_file.PnVa_data)
            output1.start_output()
            output1.write_base_data(excel_input_file.output_lista_cegiel,
                                    excel_input_file.output_leki,
                                    excel_input_file.daty)
            output1.write_PnVa_UNITS_data()
            output1.save_output(self.PnVa_file)
            print "Zapisywanie danych Zakończone OK"

            print "START zapis danych UNITS do pliku output"
            output2 = Write_output_PnVa_Units_paternA(excel_input_file.output_zakladki,
                                                      excel_input_file.UNITS_data)
            output2.start_output()
            output2.write_base_data(excel_input_file.output_lista_cegiel,
                                    excel_input_file.output_leki,
                                    excel_input_file.daty)
            output2.write_PnVa_UNITS_data()
            output2.save_output(self.UNITS_file)
            print "Zapisywanie danych Zakończone OK"

            print "START zapis danych CEGLY do pliku output"

            output3 = Write_output_CEGLY_paternA()
            output3.start_output(excel_input_file.output_lista_cegiel)
            output3.write_base_data(excel_input_file.output_lista_cegiel,
                                    excel_input_file.output_leki_cegly,
                                    excel_input_file.daty)
            print 40 * "%"
            print "excel_input_file.CEGLY_data", excel_input_file.CEGLY_data

            output3.write_PnVa_UNITS_data(excel_input_file.output_lista_cegiel,
                                          excel_input_file. output_leki_cegly,
                                          excel_input_file.CEGLY_data)
            output3.save_output(self.CEGLY_file)
            print "Zapisywanie danych Zakończone OK"


            return self.work_finished()

    # Main GUI
    #########################################3

    def interface_elem(self):
        """ Methods for creating main GUI."""

        self.get_tables_from_db()                  # Reading profile from sqlite

        # intro
        l_intro = Tkinter.Label(self.top, text=self.intro, relief="ridge", pady=2, padx=400)

        b_instruction = Tkinter.Button(self.top, text="Instrukcja obsługi programu", borderwidth=2, bg="orange",
                                       command=InformacjaTop, pady=5, padx=20)
        # ==========================================================
        buttons = Tkinter.Frame(self.top).grid(row=1, column=0, columnspan=9, sticky="nswe")

        b_profile_menager = Tkinter.Button(buttons, text="Otwórz Menadżer Profili", padx=30, pady=10,
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
        self.progres = ttk.Progressbar(self.top, orient="horizontal", mode='determinate', maximum=100, length=250)

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
        self.progres.grid(row=10, column=3, columnspan=1)

        empty_line6 = Tkinter.Label(self.top).grid(row=11, column=0, columnspan=5)

        l_licence_info.grid(row=12, column=0, columnspan=2, sticky=("we"))
        b_close.grid(row=12, column=5, sticky="nswe")

        for x in range(9):
            self.top.grid_columnconfigure(x,weight=1)

        return self.progres, self.com_choosen_profile

if __name__ == '__main__':

    input_file_path2 = "/home/marcin/Pulpit/MyProjectGitHub/report2.xlsx"
    output_file_PnVa = "/home/marcin/Pulpit/a.xlsx"
    output_file_UNITS = "/home/marcin/Pulpit/b.xlsx"
    output_file_CEGLY = "/home/marcin/Pulpit/c.xlsx"

    maingui = MainGui()
    maingui.interface_elem()







