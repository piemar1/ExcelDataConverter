# ! /usr/bin/env python
# -*- coding: utf-8 -*-
__author__ = 'Marcin Pieczyński'


import openpyxl
from edc_sqlite import SQliteEdit
import time

input_file_path1 = "/home/marcin/Pulpit/ExcelDataConverter_2.1/Raporty 3V2015/A DYMOW SZEMBEK SZYMEL Konkurencja Rx 2015_1Q.xlsx"
input_file_path2 = "/home/marcin/Pulpit/ExcelDataConverter_2.1/Raporty 3V2015/A4 Konkurencja Rx 2014_4Q.xlsx"


class Read_input(SQliteEdit):
    """Klasa zawierająca metody czytania danych z pliku excel input"""

    def __init__(self, input_file_path):
        """ Inicjalizuje obiekt pliku input - jako plik excel. """
        SQliteEdit.__init__(self)

        self.get_data_from_profile('ProfilTestowy1')     # pobieranie danych dla profilu TOP TRZEBA przenieść !!!!!!!!

        self.alfabet = "ABCDEFGHIJKLMNOPRSTUWYZ"
        self.input_file_path = input_file_path
        self.wb = None
        self.wb_sheets = None
        self.PnVa = None              # nazwa pierwszej kolumny z wynikami PnVa
        self.UNITS = None             # nazwa pierwszej kolumny z wynikami UNITS

        self.daty = []                # daty z pliku input
        self.cegly = []               # lista z nazwami cegieł z pliku input excel
        self.cegly_position = {}      # położenie danych dla poszczególnych cegieł w zakłądkach input

        self.open_input_file()
        self.get_date()
        self.get_PnVa_Units_column()
        self.get_cegla_positions()

        # najważniejsze  - dane dla poszczególnych leków
        self.PnVa_data = {}           # dane PnVa dla wybranych leków i cegieł
        self.UNITS_data = {}          # dane UNITS dla wybranych leków i cegieł
        self.CEGLY_data = {}          # dane PnVa oraz UNITS dla wybranych leków i cegieł

        # self.cegly_position = {sheet_name:{"cegla1":[10,20], "cegla2":[20,30],...kolejne cegły..},
        #                        sheet_name:{"cegla1":[10,20], "cegla2":[20,30],...kolejne cegły..},
        #                       ...kolejnr shhety...}

        # self.PnVa_data = {nazwa grupy leków1: {lek1: {cegła1: wartość, cegła2: wartość, cegła3: wartość,
        #                                              cegła4: wartość, cegła5: wartość, cegła5: wartość,}
        #                                      {lek2: {cegła1: wartość, cegła2: wartość, cegła3: wartość,
        #                                              cegła4: wartość, cegła5: wartość, cegła5: wartość,}
        #                 {nazwa grupy leków2: {lek1: {cegła1: wartość, cegła2: wartość, cegła3: wartość,
        #                                              cegła4: wartość, cegła5: wartość, cegła5: wartość,}
        #                                      {lek2: {cegła1: wartość, cegła2: wartość, cegła3: wartość,
        #                                              cegła4: wartość, cegła5: wartość, cegła5: wartość,}}}
        # self.CEGLY_data = {cegła1: {lek1: {PnVa: [3 x wartość], UNITS: [3xwartość]},
        #                            {lek2: {PnVa: [3 x wartość], UNITS: [3xwartość]},
        #                   {cegła2: {lek1: {PnVa: [3 x wartość], UNITS: [3xwartość]},
        #                            {lek2: {PnVa: [3 x wartość], UNITS: [3xwartość]},

    def open_input_file(self):
        """ Metoda otwiera plik excel oraz zczytuje zakładki po nazwach. """
        self.wb = openpyxl.load_workbook(self.input_file_path)
        self.wb_sheets = self.wb.get_sheet_names()
        # print self.wb_sheets

    def get_date(self):
        """ Metoda odnajduje w pliku intup informacje o dacie zawartych danych. """
        for sheet_name in list(self.wb_sheets):
            zakladka = self.wb.get_sheet_by_name(sheet_name)
            size = zakladka.calculate_dimension() # zwraca wielkość zakładki w postaci strinha --> A1:T60
            for row in zakladka.iter_rows(size):
                for cell in row:
                    if "MNTH" in str(cell.value):
                        d = str(cell.value)[-7:]
                        if d not in self.daty:
                            self.daty.append(d)
                        if len(self.daty) == 3:
                            break

    def get_PnVa_Units_column(self):
        """ Metoda odnajduje kolumnmy zawierające dane PnVa oraz UNITS. """
        for sheet_name in list(self.wb_sheets):
            zakladka = self.wb.get_sheet_by_name(sheet_name)
            size = zakladka.calculate_dimension() # zwraca wielkość zakładki w postaci strinha --> A1:T60
            for row in zakladka.iter_rows(size):
                for cell in row:
                    if "Pn Va" in str(cell.value):
                        self.PnVa = cell.column
                    if "UNITS REPORT" in str(cell.value):
                        self.UNITS = cell.column
                    if self.PnVa and self.UNITS:
                        break
        # print ("PnVa", self.PnVa)
        # print ("Units", self.UNITS)

    def get_cegla_positions(self):
        """ Metoda znajduje i zbiera informacje na temat cegieł zawartym w pliku excel input"""

        for sheet_name in list(self.wb_sheets):
            zakladka = self.wb.get_sheet_by_name(sheet_name)
            no_of_row = zakladka.get_highest_row()              # zwraca liczbę row
            for r in range(no_of_row):                          # iteracja po row
                wart = zakladka["A" + str(r+1)].value
                if wart and len(wart)<= 13:
                    # BARDZO NEWRALGICZNE miejsce liczba 13 - uzależniona od liczby spacji przy nazwach cegieł
                    # w pliku excel !!!!!!!!!!!!
                    if wart.strip() not in self.cegly:
                        self.cegly.append(wart.strip())

        for sheet_name in list(self.wb_sheets):
            positions = []
            self.cegly_position[sheet_name] = {str(k): [None, None] for k in self.cegly}
            zakladka = self.wb.get_sheet_by_name(sheet_name)
            no_of_row = zakladka.get_highest_row()                  # zwraca liczbę row

            for cegla in self.cegly:
                for r in range(no_of_row):                          # iteracja po row
                    wart = zakladka["A" + str(r+1)].value
                    if wart and wart.strip() == cegla:
                        positions.append(r+1)

            # print  "positions", positions
            for x in range(len(self.cegly)-1):              # zapisywanie w słowniku informacji o położeniu cegieł
                self.cegly_position[sheet_name][self.cegly[x]][0] = positions[x]
                self.cegly_position[sheet_name][self.cegly[x]][1] = positions[x+1]-1

            self.cegly_position[sheet_name][self.cegly[-1]][0] = positions[-1]
            self.cegly_position[sheet_name][self.cegly[-1]][1] = int(no_of_row)    # To miejsce można zooptymalizować
            # int(no_of_row) często wykracza znacząco poza rzeczywistą wielkość tabeli

    def get_Cegly_data(self):
        """ Metoda wypełnia danymi z pliku input słownik dla danych PnVa oraz UNITS dla cegiel"""


    # Tworzenie dużego słownika dla danych PnVa
        for cegla in self.cegly:
            self.CEGLY_data[cegla] = {lek: {"PnVa": [None, None, None], "UNITS": [None, None, None]}
                                      for lek in self.output_leki_cegly}

    # UZUPEŁANIANIE dużego słownika dla danych PnVa i UNITS w CEGLACH
        hit_number = len(self.output_leki_cegly)
        column_PnVa_plus1 = self.alfabet[self.alfabet.index(self.PnVa)+1]
        column_PnVa_plus2 = self.alfabet[self.alfabet.index(self.PnVa)+2]
        column_UNITS_plus1 = self.alfabet[self.alfabet.index(self.UNITS)+1]
        column_UNITS_plus2 = self.alfabet[self.alfabet.index(self.UNITS)+2]

        for sheet_name in self.wb_sheets:
            zakladka = self.wb.get_sheet_by_name(sheet_name)
            size = zakladka.calculate_dimension()        # zwraca wielkość zakładki w postaci stringa --> A1:T60
            for row in zakladka.iter_rows(size):
                for cell in row:
                    for lek in self.output_leki_cegly:
                        if lek in str(cell.value):
                            hit_number -= 1
                            wart1 = zakladka[str(self.PnVa) + str(cell.row)].value
                            wart2 = zakladka[str(column_PnVa_plus1) + str(cell.row)].value
                            wart3 = zakladka[str(column_PnVa_plus2) + str(cell.row)].value
                            wart4 = zakladka[str(self.UNITS) + str(cell.row)].value
                            wart5 = zakladka[str(column_UNITS_plus1) + str(cell.row)].value
                            wart6 = zakladka[str(column_UNITS_plus2) + str(cell.row)].value
                            for cegla in self.cegly:
                                x = self.cegly_position[sheet_name][cegla][0]
                                y = self.cegly_position[sheet_name][cegla][1]
                                if cell.row >= x and cell.row <= y:
                                    self.CEGLY_data[cegla][lek]["PnVa"] = [round(wart1, 2),
                                                                           round(wart2, 2),
                                                                           round(wart3, 2)]
                                    self.CEGLY_data[cegla][lek]["UNITS"] = [round(wart4, 2),
                                                                            round(wart5, 2),
                                                                            round(wart6, 2)]
                                    if hit_number == 0:
                                        break

        print "\nzawartość słownika --> self.CEGLY_data\n"
        for k, v in self.CEGLY_data.iteritems():
            print 5 * "xx"
            print k, v

    def get_PnVa_UNITS_data(self):
        """ Metoda wypełnia danymi z pliku input dwa słowniki odpowiednio dla danych PnVa oraz UNITS"""

        self.get_data_from_profile('ProfilTestowy1')     # pobieranie danych dla profilu TOP TRZEBA przenieść !!!!!!!!

    # Tworzenie dużego słownika dla danych PnVa
        for no_x in range(len(self.output_zakladki)):
                self.PnVa_data[self.output_zakladki[no_x]] = \
                    {str(k): {cegla: [None, None, None] for cegla in self.cegly} for k in self.output_leki[no_x]}

    # Tworzenie dużego słownika dla danych UNITS
        for no_x in range(len(self.output_zakladki)):
                self.UNITS_data[self.output_zakladki[no_x]] = \
                    {str(k): {cegla: [None, None, None] for cegla in self.cegly} for k in self.output_leki[no_x]}

    # Uzupełnianie słownika PnVa z danym z pliku input !!!!!!!!!!!!!!!!
        [self.uzupelnianie_tabeli_PnVa_UNITS(self.output_leki[elem], self.output_zakladki[elem], self.PnVa)
         for elem in range(len(self.output_zakladki))]            # a może by zastosować funkcję map......

    # Uzupełnianie słownika UNITS z danym z pliku input !!!!!!!!!!!!!!!!
        [self.uzupelnianie_tabeli_PnVa_UNITS(self.output_leki[elem], self.output_zakladki[elem], self.UNITS)
         for elem in range(len(self.output_zakladki))]            # a może by zastosować funkcję map......

        # print "\nzawartość słownika --> self.PnVa_data\n"
        # for k, v in self.PnVa_data.iteritems():
        #     print 5 * "xx"
        #     print k, v
        #
        # print "\nzawartość słownika --> self.UNITS_data\n"
        # for k, v in self.UNITS_data.iteritems():
        #     print 5 * "xx"
        #     print k, v

    # UZUPEŁANIANIE dużego słownika dla danych PnVa lub UNITS
    def uzupelnianie_tabeli_PnVa_UNITS(self, lista_lekow, nazwa_zakladki, rodzaj_danych):
        """ Funckcja zbiera wartości z pliku input i zapisuje je w tabeli / słowniku PnVa UNITS"""

        if rodzaj_danych == self.PnVa:
            column_plus1 = self.alfabet[self.alfabet.index(self.PnVa)+1]
            column_plus2 = self.alfabet[self.alfabet.index(self.PnVa)+2]
            slownik = self.PnVa_data
        elif rodzaj_danych == self.UNITS:
            column_plus1 = self.alfabet[self.alfabet.index(self.UNITS)+1]
            column_plus2 = self.alfabet[self.alfabet.index(self.UNITS)+2]
            slownik = self.UNITS_data

        hit_number = len(lista_lekow)

        for sheet_name in self.wb_sheets:
            zakladka = self.wb.get_sheet_by_name(sheet_name)
            size = zakladka.calculate_dimension()        # zwraca wielkość zakładki w postaci stringa --> A1:T60
            for row in zakladka.iter_rows(size):
                for cell in row:
                    for lek in lista_lekow:
                        if lek in str(cell.value):
                            hit_number -= 1
                            wart1 = zakladka[str(rodzaj_danych) + str(cell.row)].value
                            wart2 = zakladka[str(column_plus1) + str(cell.row)].value
                            wart3 = zakladka[str(column_plus2) + str(cell.row)].value
                            for cegla in self.cegly:
                                x = self.cegly_position[sheet_name][cegla][0]
                                y = self.cegly_position[sheet_name][cegla][1]
                                if cell.row >= x and cell.row <= y:
                                    slownik[nazwa_zakladki][lek][cegla][0] = round(wart1, 2)
                                    slownik[nazwa_zakladki][lek][cegla][1] = round(wart2, 2)
                                    slownik[nazwa_zakladki][lek][cegla][2] = round(wart3, 2)
                                    if hit_number == 0:
                                        break


if __name__ == '__main__':



    input1 = Read_input(input_file_path1)
    print "input1"
    print input1.cegly
    # print input1.daty
    # print "zawartość słownika --> self.cegly_position"
    # for k, v in input1.cegly_position.iteritems():
    #     print k, v
    print 50 * "xxx"

    input1.get_Cegly_data()


    # input1.get_PnVa_UNITS_data()
    # print input1.PnVa_data
    print 50 * "%%%"


    print "output_zakladki", self.output_zakladki


