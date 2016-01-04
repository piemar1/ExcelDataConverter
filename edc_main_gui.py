# ! /usr/bin/env python
# -*- coding: utf-8 -*-

import Tkinter
import ttk
import tkMessageBox
import os
import tkFileDialog

from edc_sqlite import SQliteEdit
from edc_init import LicenceCheck
from edc_read_input import ReadInput
from edc_write_output import WriteOutputPnVaUnitsPaternA, WriteOutputCeglyPaternA
import edc_profile_manager

from time import time

__author__ = 'Marcin Pieczyński'

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
    The window with information about usage the program.
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
    """
    The Class containing main gui of the program and all methods connected with the gui.
    """
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

        self.imputnumber = 0
        self.PvNanumber = 0
        self.UNITSnumber = 0
        self.Ceglynumber = 0
        self.progress_step = 0
        self.filepath_input = Tkinter.StringVar()
        self.filepath_PnVa = Tkinter.StringVar()
        self.filepath_UNITS = Tkinter.StringVar()
        self.filepath_CEGLY = Tkinter.StringVar()

        self.licence_days = "Do końca licencji pozostało {} dni.".format(str(self.delta))

        self.excel_input_file = None
        self.PnVa_file = None
        self.UNITS_file = None
        self.CEGLY_file = None
        self.com_choosen_profile = None
        self.progress = None

        self.interface_elem()
        self.top.mainloop()

    def start_progress_bar(self):
        """
        Start of the progress bar.
        """
        self.progress["value"] = self.progress_step

    def ad_step_to_progress_bar(self, n):
        """
        Update of the progress bar.
        """
        self.progress_step += n
        self.progress["value"] = self.progress_step
        self.progress.update_idletasks()

    def input_file(self):
        """
        The method for introduction an input file with data for filtration.
        """
        try:
            f = tkFileDialog.askopenfilename(parent=self.top, initialdir="/home/marcin/pulpit/Py/",
                                             title="Wybór pliku excel z danymi",
                                             filetypes=[("Excel file", ".xlsx")])
            self.filepath_input.set(os.path.realpath(f))
            self.excel_input_file = os.path.realpath(f)
            self.imputnumber += 1
        except ValueError:
            tkMessageBox.showerror("Error", "Wystąpił problem z załadowaniem pliku excel z danymi.")

    def save_file_pnva(self):
        """
        The method for introduction an output file with PnVa data.
        """
        try:
            save = tkFileDialog.asksaveasfilename(parent=self.top, initialdir="/home/marcin/pulpit/",
                                                  title="Wybór pliku do zapisu danych PnVa",
                                                  filetypes=[("Excel file", ".xlsx")])
            self.filepath_PnVa.set((os.path.realpath(save)))
            self.PnVa_file = os.path.realpath(save)
            self.PvNanumber += 1
        except ValueError:
            tkMessageBox.showerror("Error", " Wystąpił problem z plikiem do zapisu danych PnVa.")

    def save_file_units(self):
        """
        The method for introduction an output file with UNITS data.
        """
        try:
            save = tkFileDialog.asksaveasfilename(parent=self.top, initialdir="/home/marcin/pulpit/",
                                                  title="Wybór pliku do zapisu danych UNITS",
                                                  filetypes=[("Excel file", ".xlsx")])
            self.filepath_UNITS.set((os.path.realpath(save)))
            self.UNITS_file = os.path.realpath(save)
            self.UNITSnumber += 1
        except ValueError:
            tkMessageBox.showerror("Error", " Wystąpił problem z plikiem do zapisu danych UNITS.")

    def save_file_cegly(self):
        """
        The method for introduction an output file with area data.
        """
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
        """
        The methods launching profile manager.
        """
        reload(edc_profile_manager)
        edc_profile_manager.ProfileManager()

    def error_no_profile(self):
        """
        The window with message informing that the profile has not been chosen..
        """
        tkMessageBox.showerror("Błąd, Nie wybrano profilu.",
                               "Należy dokonać wyboru profilu do filtrowania danych, "
                               "szczegóły poszczególnych profili można znaleść w Menadżeże Profili")

    def error_no_files(self):
        """
        The window with message informing that the files have not been chosen.
        """
        tkMessageBox.showerror(" Błąd, Nie wybrano nazw plików. ",
                               "Przed konwersją danych należy wybrać plik wejściowy zawierające dane Excel "
                               "oraz podać nazwy dla 3 plików wyjściowych w których program zapisze dane.")

    def work_finished(self):
        """
        The window with message informing that excel data filtration was finished.
        """
        tkMessageBox.showinfo("Yes...", "Dokonano konwersji danych. \n Życzę miłego dnia.")

    def krka_data_convert(self):
        """ Important !  The methods for carrying out reading, filtration and saving of excel data. """

        profil = self.com_choosen_profile.get()

        if profil == "Wybierz profil":
            return self.error_no_profile()

        elif self.imputnumber == 0 or self.PvNanumber == 0 or self.UNITSnumber == 0 or self.Ceglynumber == 0:
            return self.error_no_files()

        elif self.imputnumber and self.PvNanumber and self.UNITSnumber and self.Ceglynumber \
                and profil != "Wybierz profil":

            ###############################################################
            # Methods for for carrying out reading, filtration and saving of excel data.

            start = time()
            self.start_progress_bar()
            print "1", time() - start

            # Reading of input data for excel file.

            excel_input_file = ReadInput(profil)
            print "2", time() - start

            excel_input_file.open_input_file(self.excel_input_file)
            self.ad_step_to_progress_bar(5)
            print "3", time() - start
            excel_input_file.get_date()
            self.ad_step_to_progress_bar(5)
            print "4", time() - start
            excel_input_file.get_pnva_units_column()
            self.ad_step_to_progress_bar(5)
            print "5", time() - start
            excel_input_file.get_cegla_positions()
            self.ad_step_to_progress_bar(5)

            print "6", time() - start
            excel_input_file.get_cegly_data()
            self.ad_step_to_progress_bar(5)

            print "7", time() - start
            excel_input_file.get_pnva_units_data()
            self.ad_step_to_progress_bar(5)
            # print "Odczytanie danych z pliku input OK"

            print "8", time() - start
            # Saving PnVa data into output file
            print "9", time() - start
            # print "START zapis danych PnVa do pliku output"
            output1 = WriteOutputPnVaUnitsPaternA(excel_input_file.output_zakladki,
                                                  excel_input_file.PnVa_data)
            print "10", time() - start
            self.ad_step_to_progress_bar(5)
            print "11", time() - start
            output1.start_output()
            self.ad_step_to_progress_bar(5)
            print "12", time() - start
            output1.write_base_data(excel_input_file.output_lista_cegiel,
                                    excel_input_file.output_leki,
                                    excel_input_file.daty)
            print "13", time() - start
            self.ad_step_to_progress_bar(5)
            print "14", time() - start
            output1.write_pnva_units_data()
            print "15", time() - start
            self.ad_step_to_progress_bar(5)
            print "16", time() - start
            output1.save_output(self.PnVa_file)
            print "17", time() - start
            self.ad_step_to_progress_bar(5)
            print "18", time() - start

            # print "Zapisywanie danych Zakończone OK"

            # Saving UNITS data into output file

            # print "START zapis danych UNITS do pliku output"

            output2 = WriteOutputPnVaUnitsPaternA(excel_input_file.output_zakladki,
                                                  excel_input_file.UNITS_data)
            print "19", time() - start
            self.ad_step_to_progress_bar(5)
            print "20", time() - start
            output2.start_output()
            self.ad_step_to_progress_bar(5)
            print "21", time() - start
            output2.write_base_data(excel_input_file.output_lista_cegiel,
                                    excel_input_file.output_leki,
                                    excel_input_file.daty)
            print "22", time() - start
            self.ad_step_to_progress_bar(5)
            print "23", time() - start
            output2.write_pnva_units_data()
            print "24", time() - start
            self.ad_step_to_progress_bar(5)
            print "25", time() - start
            output2.save_output(self.UNITS_file)
            print "26", time() - start
            self.ad_step_to_progress_bar(5)
            print "27", time() - start

            # print "Zapisywanie danych Zakończone OK"

            # Saving CEGLY data into output file

            # print "START zapis danych CEGLY do pliku output"

            output3 = WriteOutputCeglyPaternA()
            print "28", time() - start
            self.ad_step_to_progress_bar(5)
            print "29", time() - start
            output3.start_output(excel_input_file.output_lista_cegiel)
            print "30", time() - start
            self.ad_step_to_progress_bar(5)
            print "31", time() - start
            output3.write_base_data(excel_input_file.output_lista_cegiel,
                                    excel_input_file.output_leki_cegly,
                                    excel_input_file.daty)
            print "32", time() - start
            self.ad_step_to_progress_bar(5)
            print "33", time() - start
            output3.write_pnva_units_data(excel_input_file.output_lista_cegiel,
                                          excel_input_file. output_leki_cegly,
                                          excel_input_file.CEGLY_data)
            print "34", time() - start
            self.ad_step_to_progress_bar(5)
            print "35", time() - start
            output3.save_output(self.CEGLY_file)
            print "36", time() - start
            self.ad_step_to_progress_bar(5)
            print "37", time() - start

            print "Zapisywanie danych Zakończone OK"

            return self.work_finished()

    #####
    # Main GUI

    def interface_elem(self):
        """
        Methods for creating main GUI.
        """

        self.get_tables_from_db()                  # Reading profile from sqlite

        # intro
        l_intro = Tkinter.Label(self.top, text=self.intro, relief="ridge", pady=2, padx=400)

        b_instruction = Tkinter.Button(self.top, text="Instrukcja obsługi programu", borderwidth=2, bg="orange",
                                       command=InformacjaTop, pady=5, padx=20)
        ###
        buttons = Tkinter.Frame(self.top).grid(row=1, column=0, columnspan=9, sticky="nswe")

        b_profile_menager = Tkinter.Button(buttons, text="Otwórz Menadżer Profili", padx=30, pady=10,
                                           command=self.open_profile_menager)

        b_profile_reload = Tkinter.Button(buttons, text="Odśwież listę profili",
                                          padx=30, pady=10, command=self.interface_elem)

        self.com_choosen_profile = ttk.Combobox(buttons)
        self.com_choosen_profile.insert('0', "Wybierz profil")
        self.com_choosen_profile['values'] = self.profiles_name_list

        ###
        #  input file
        b_input_file = Tkinter.Button(self.top, text="Plik wejsciowy", command=self.input_file, padx=60)
        l_input_file = Tkinter.Label(self.top, width=7, textvariable=self.filepath_input)

        # output file PvNa
        b_output_pnva = Tkinter.Button(self.top, text="Plik wyjściowy   PnVa", command=self.save_file_pnva)
        l_output_pnva = Tkinter.Label(self.top, width=7, textvariable=self.filepath_PnVa)

        # output file UNITS
        b_output_units = Tkinter.Button(self.top, text="Plik wyjściowy   UNITS", command=self.save_file_units)
        l_output_units = Tkinter.Label(self.top, width=7, textvariable=self.filepath_UNITS)

        # output file CEGŁY
        b_output_cegly = Tkinter.Button(self.top, text="Plik wyjściowy   CEGŁY", command=self.save_file_cegly)
        l_output_cegly = Tkinter.Label(self.top, width=7, textvariable=self.filepath_CEGLY)

        # button converting data !!!
        b_convert = Tkinter.Button(self.top, text="Konwertuj dane !!!", command=self.krka_data_convert)
        self.progress = ttk.Progressbar(self.top, orient="horizontal", mode='determinate', maximum=105, length=250)

        # the number of days until the license expires
        l_licence_info = Tkinter.Label(self.top, width=1, text=self.licence_days)

        # close the program
        b_close = Tkinter.Button(self.top, text="Zamknij program", command=self.top.destroy,
                                 borderwidth=2, relief="ridge", pady=4, padx=30)

        # Localization of GUI elements
        l_intro.grid(row=0, column=0, columnspan=5, sticky="nswe")
        b_instruction.grid(row=0, column=5, sticky="nswe")

        b_profile_menager.grid(row=1, column=1, sticky="we")
        b_profile_reload.grid(row=1, column=3, sticky="we")
        self.com_choosen_profile.grid(row=1, column=4, sticky="we", padx=50, pady=10, ipadx=20)
        l_empty_line = Tkinter.Label(buttons).grid(row=1, column=2)
        l_empty_line2 = Tkinter.Label(buttons).grid(row=2, column=0, columnspan=5)

        b_input_file.grid(row=2, column=0, sticky="nswe")
        l_input_file.grid(row=2, column=1, columnspan=10, sticky="we")
        l_empty_line3 = Tkinter.Label(self.top).grid(row=3, column=0, columnspan=5)

        b_output_pnva.grid(row=4, column=0, sticky="nswe")
        l_output_pnva.grid(row=4, column=1, columnspan=10, sticky="we")
        b_output_units.grid(row=5, column=0, sticky="nswe")
        l_output_units.grid(row=5, column=1, columnspan=10, sticky="we")
        b_output_cegly.grid(row=6, column=0, sticky="nswe")
        l_output_cegly.grid(row=6, column=1, columnspan=10, sticky="we")

        l_empty_line4 = Tkinter.Label(self.top).grid(row=7, column=0, columnspan=5)
        b_convert.grid(row=8, column=3, sticky="nswe")

        l_empty_line5 = Tkinter.Label(self.top).grid(row=9, column=0, columnspan=5)
        self.progress.grid(row=10, column=3, columnspan=1)

        empty_line6 = Tkinter.Label(self.top).grid(row=11, column=0, columnspan=5)

        l_licence_info.grid(row=12, column=0, columnspan=2, sticky="we")
        b_close.grid(row=12, column=5, sticky="nswe")

        for x in range(9):
            self.top.grid_columnconfigure(x, weight=1)

        return self.progress, self.com_choosen_profile

if __name__ == '__main__':
    maingui = MainGui()
    maingui.interface_elem()
