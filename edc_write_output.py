# ! /usr/bin/env python
# -*- coding: utf-8 -*-

import openpyxl
from edc_read_input import ReadInput
from edc_sqlite import SQliteEdit
from string import ascii_uppercase

__author__ = 'Marcin Pieczyński'


class WriteOutputPnVaUnitsPaternA(object):
    """
    Class containing methods for writing PnVa and UNITS data to excel output file.
    """
    def __init__(self, output_zakladki, slownik):

        self.alfabet = ascii_uppercase
        self.output_zakladki = output_zakladki
        self.slownik = slownik

        self.output_sheets = None

    def start_output(self):
        """
        Method for creating output file with sheets.
        """
        self.output_file = openpyxl.Workbook()

        for elem in range(len(self.output_zakladki)):
            self.output_file.create_sheet(self.output_zakladki[elem], elem)

        # Delete first sheet with defaults name
        self.output_file.remove_sheet(self.output_file.get_sheet_by_name('Sheet'))
        self.output_sheets = self.output_file.get_sheet_names()

    def save_output(self, output_file_path):
        """
        Methods for saving excel output file in a given path.
        """
        self.output_file.save(output_file_path)

    def write_base_data(self, output_lista_cegiel, output_leki, daty):
        """
        Zapisuje w zakladkach podstawowe dane (daty, cegły i nazwy leków.
        """
        self.output_lista_cegiel = output_lista_cegiel
        self.output_leki = output_leki
        self.daty = daty

        ################################################
        # rozdzielić te finkcje na 3 mniejsze !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!1111
        # ######################################################################3

        liczba_pelnych_linii = len(self.output_lista_cegiel) / 5
        if len(self.output_lista_cegiel)-liczba_pelnych_linii * 5 != 0:
            niepelna_liniia = True
        else:
            niepelna_liniia = False

        # zapisywanie danych leków w pliku output w poszczególnych zakładkach
        for elem in range(len(self.output_zakladki)):
            n, first_row = 0, 3

            sheet = self.output_zakladki[elem]
            sh = self.output_file.get_sheet_by_name(sheet)

            no_of_row = liczba_pelnych_linii
            if niepelna_liniia:
                no_of_row += 1

            while no_of_row != 0:
                for lek in self.output_leki[elem]:
                    sh['A' + str(first_row + n)] = lek
                    n += 1
                first_row += 2
                no_of_row -= 1

        # zapisywanie nazw cegiel w pliku output w poszczególnych zakładkach
        for elem in range(len(self.output_zakladki)):
            first_row, first_col = 1, 3

            sheet = self.output_zakladki[elem]
            sh = self.output_file.get_sheet_by_name(sheet)

            cegla_no = 0
            for no in range(liczba_pelnych_linii):                     # drukuje pełne linie liczące po 5 cegieł
                for i in range(0, 5):
                    sh[self.alfabet[first_col + i] + str(first_row)] = self.output_lista_cegiel[cegla_no]
                    first_col += 2
                    cegla_no += 1
                first_col = 3
                first_row += len(self.output_leki[elem]) + 2

            # print len(self.output_lista_cegiel), cegla_no
            for i in range(len(self.output_lista_cegiel) - cegla_no):                 # drukuje NIEpełne linie
                sh[self.alfabet[first_col] + str(first_row)] = self.output_lista_cegiel[cegla_no]
                cegla_no += 1
                first_col += 3

        # zapisywanie dat w pliku output w poszczególnych zakładkach
        for elem in range(len(self.output_zakladki)):
            first_row, first_col = 2, 1

            sheet = self.output_zakladki[elem]
            sh = self.output_file.get_sheet_by_name(sheet)
            no_date, cegla_no = 0, 0

            for no in range(liczba_pelnych_linii):      # drukuje pełne linie liczące po 5 cegieł
                for i in range(15):
                    sh[self.alfabet[first_col] + str(first_row)] = self.daty[no_date]
                    first_col += 1
                    no_date += 1
                    if no_date == 3:
                        no_date = 0
                cegla_no += 5
                first_col = 1
                first_row += len(self.output_leki[elem]) + 2

            # print len(self.output_lista_cegiel), cegla_no
            mising_date = (len(self.output_lista_cegiel) - cegla_no) * 3
            for i in range(mising_date):                                     # drukuje NIEpełne linie
                sh[self.alfabet[first_col] + str(first_row)] = self.daty[no_date]
                first_col += 1
                no_date += 1
                if no_date == 3:
                    no_date = 0

    def write_pnva_units_data(self):
        """
        Method for writing in side output file PnVa and UNITS data.
        """
        for elem in range(len(self.output_zakladki)):
            cegla_no, col, row = 0, 1, 3
            no_of_row = len(self.output_lista_cegiel) / 5

            sheet = self.output_zakladki[elem]
            sh = self.output_file.get_sheet_by_name(sheet)

            for e in range(no_of_row):                    # writing data in complete row, containing 5 cegla
                for lek in self.output_leki[elem]:
                    for i in range(5):
                        for ii in range(3):
                            wart = self.slownik[sheet][lek][self.output_lista_cegiel[i]][ii]
                            sh[self.alfabet[col] + str(row)] = wart
                            col += 1
                        cegla_no += 1
                    col = 1
                    row += 1
                row += 2

            mising_cegla = (len(self.output_lista_cegiel) - (cegla_no / len(self.output_leki[elem])))
            tru_cegla_no = (cegla_no / len(self.output_leki[elem]))

            for lek in self.output_leki[elem]:            # writing data in incomplete row, containing < 5 cegla
                for i in range(mising_cegla):
                    for ii in range(3):
                        wart = self.slownik[sheet][lek][self.output_lista_cegiel[tru_cegla_no]][ii]
                        sh[self.alfabet[col] + str(row)] = wart
                        col += 1
                    tru_cegla_no += 1
                col = 1
                row += 1
                tru_cegla_no = (cegla_no / len(self.output_leki[elem]))


class WriteOutputCeglyPaternA(object):
    """
    Class containing methods for writing cegla data to excel output file.
    """
    def __init__(self):
        self.alfabet = ascii_uppercase
        self.output_file = None

    def start_output(self, output_lista_cegiel):
        """
        Method for creating output file with sheets.
        """
        self.output_file = openpyxl.Workbook()
        for elem in range(len(output_lista_cegiel)):
            self.output_file.create_sheet(output_lista_cegiel[elem], elem)

        # Delete first sheet with defaults name
        self.output_file.remove_sheet(self.output_file.get_sheet_by_name('Sheet'))
        self.output_sheets = self.output_file.get_sheet_names()

    def save_output(self, output_file_path):
        """
        Methods for saving excel output file in a given path.
        """
        self.output_file.save(output_file_path)

    def write_base_data(self, output_lista_cegiel, output_leki_cegly, daty):
        """
        Methods for writing base data to output file with cegla data.
        """
        for elem in range(len(output_lista_cegiel)):
            sheet = output_lista_cegiel[elem]
            sh = self.output_file.get_sheet_by_name(sheet)
            col = 1
            sh['B3'] = "UNITS"
            sh['E3'] = "PnVa"

            for i in range(2):
                for ii in range(3):
                    sh[self.alfabet[col] + '4'] = daty[ii]
                    col += 1
            col, row = 0, 5
            for i in output_leki_cegly:
                sh[self.alfabet[col] + str(row)] = i
                row += 1

    def write_pnva_units_data(self, output_lista_cegiel, output_leki_cegly, CEGLY_data):
        """
        Methods for writing PnVa and UNITS data to output file with UNITS data.
        """
        for elem in range(len(output_lista_cegiel)):
            col, row = 1, 5
            sheet = output_lista_cegiel[elem]
            sh = self.output_file.get_sheet_by_name(sheet)

            for i in output_leki_cegly:
                for x in range(3):
                    sh[self.alfabet[col] + str(row)] = CEGLY_data[sheet][i]["UNITS"][x]
                    col += 1
                for x in range(3):
                    sh[self.alfabet[col] + str(row)] = CEGLY_data[sheet][i]["PnVa"][x]
                    col += 1
                col = 1
                row += 1

if __name__ == '__main__':

    input_file_path1 = "/home/marcin/Pulpit/MyProjectGitHub/report1.xlsx"
    input_file_path2 = "/home/marcin/Pulpit/MyProjectGitHub/report2.xlsx"

    output_zakladki = [u'roswera', u'atoris']

    PnVa_data = {u'roswera': {'SUVARDIO BRAND': {u'A4.06': [134.45, 180.28, 167.92], u'A4.07': [104.1, 108.98, 75.55], u'A4.04': [99.29, 95.22, 145.83], u'A4.05': [142.43, 145.96, 145.78], u'A4.02': [210.66, 169.69, 171.97], u'A4.03': [195.36, 210.2, 206.22], u'A DYMOWS SZYMEL SZEMBE': [128.17, 128.92, 129.87], u'A4.01': [198.94, 215.06, 200.03], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [133.08, 129.94, 153.28], u'A4': [144.88, 149.79, 151.97]}, 'ROMAZIC BRAND': {u'A4.06': [61.39, 117.51, 50.14], u'A4.07': [80.61, 71.2, 42.54], u'A4.04': [57.04, 89.79, 77.03], u'A4.05': [69.18, 46.87, 69.31], u'A4.02': [92.43, 74.02, 87.03], u'A4.03': [56.06, 52.76, 53.22], u'A DYMOWS SZYMEL SZEMBE': [85.24, 81.29, 82.25], u'A4.01': [73.9, 25.51, 65.94], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [119.74, 109.79, 112.42], u'A4': [81.68, 79.0, 73.46]}, 'ZAHRON BRAND': {u'A4.06': [75.03, 66.5, 74.36], u'A4.07': [93.47, 85.02, 98.7], u'A4.04': [110.41, 63.12, 106.28], u'A4.05': [71.6, 70.5, 79.07], u'A4.02': [73.39, 77.82, 113.67], u'A4.03': [108.88, 118.6, 129.11], u'A DYMOWS SZYMEL SZEMBE': [79.61, 75.7, 104.1], u'A4.01': [71.37, 106.25, 129.92], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [79.24, 104.31, 115.31], u'A4': [84.18, 87.0, 102.56]}, 'ZARANTA BRAND': {u'A4.06': [37.29, 42.73, 33.47], u'A4.07': [36.68, 38.08, 25.66], u'A4.04': [36.09, 37.46, 40.51], u'A4.05': [34.69, 25.06, 33.34], u'A4.02': [38.24, 51.73, 55.49], u'A4.03': [62.02, 64.16, 96.97], u'A DYMOWS SZYMEL SZEMBE': [76.33, 71.59, 71.91], u'A4.01': [76.53, 47.83, 37.32], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [49.08, 41.67, 50.78], u'A4': [43.86, 41.66, 46.52]}, 'ROSWERA BRAND': {u'A4.06': [153.32, 146.11, 150.42], u'A4.07': [70.02, 69.42, 81.29], u'A4.04': [98.4, 92.3, 121.76], u'A4.05': [100.96, 99.93, 109.05], u'A4.02': [107.7, 138.7, 126.84], u'A4.03': [94.16, 99.65, 97.15], u'A DYMOWS SZYMEL SZEMBE': [105.4, 105.73, 103.57], u'A4.01': [108.9, 137.62, 98.92], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [120.18, 118.63, 123.61], u'A4': [107.7, 110.04, 114.31]}}, u'atoris': {'ATORVASTEROL TOTAL': {u'A4.06': [105.11, 88.6, 89.31], u'A4.07': [72.78, 62.62, 85.01], u'A4.04': [65.67, 45.57, 70.87], u'A4.05': [72.6, 70.17, 67.04], u'A4.02': [69.48, 41.71, 46.65], u'A4.03': [65.58, 57.08, 76.96], u'A DYMOWS SZYMEL SZEMBE': [94.79, 100.6, 94.42], u'A4.01': [65.18, 63.54, 47.02], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [94.12, 92.18, 67.28], u'A4': [79.94, 70.29, 71.15]}, 'TORVACARD TOTAL': {u'A4.06': [None, None, None], u'A4.07': [None, None, None], u'A4.04': [None, None, None], u'A4.05': [None, None, None], u'A4.02': [None, None, None], u'A4.03': [None, None, None], u'A DYMOWS SZYMEL SZEMBE': [None, None, None], u'A4.01': [None, None, None], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [None, None, None], u'A4': [None, None, None]}, 'TORVALIPIN BRAND': {u'A4.06': [None, None, None], u'A4.07': [None, None, None], u'A4.04': [None, None, None], u'A4.05': [None, None, None], u'A4.02': [None, None, None], u'A4.03': [None, None, None], u'A DYMOWS SZYMEL SZEMBE': [None, None, None], u'A4.01': [None, None, None], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [None, None, None], u'A4': [None, None, None]}, 'ATORIS FILM C.TABS 30 MG 30': {u'A4.06': [60.02, 35.77, 45.12], u'A4.07': [51.72, 129.69, 11.95], u'A4.04': [62.42, 25.66, 94.85], u'A4.05': [25.51, 34.66, 22.89], u'A4.02': [11.65, 56.54, 15.32], u'A4.03': [34.29, 34.47, 30.04], u'A DYMOWS SZYMEL SZEMBE': [58.83, 58.48, 58.82], u'A4.01': [46.53, 98.0, 94.47], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [56.55, 43.63, 50.21], u'A4': [44.22, 53.97, 38.45]}, 'ATORIS FILM C.TABS 30 MG 60': {u'A4.06': [14.03, 7.88, 13.35], u'A4.07': [62.73, 74.95, 58.04], u'A4.04': [47.13, 13.21, 78.69], u'A4.05': [19.34, 51.18, 46.73], u'A4.02': [19.83, 72.13, -62.66], u'A4.03': [23.21, 41.27, 27.33], u'A DYMOWS SZYMEL SZEMBE': [42.98, 55.13, 50.63], u'A4.01': [45.15, 0.0, 64.46], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [33.23, 57.44, 44.91], u'A4': [31.61, 46.83, 34.71]}, 'TULIP TOTAL': {u'A4.06': [130.5, 131.07, 162.87], u'A4.07': [141.72, 153.79, 140.25], u'A4.04': [114.96, 128.17, 163.6], u'A4.05': [122.01, 128.06, 140.07], u'A4.02': [133.58, 134.4, 126.95], u'A4.03': [100.57, 117.41, 100.57], u'A DYMOWS SZYMEL SZEMBE': [133.94, 136.41, 135.1], u'A4.01': [130.43, 103.17, 108.91], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [114.98, 112.32, 119.8], u'A4': [122.55, 126.79, 132.8]}, 'ATORIS TOTAL': {u'A4.06': [112.16, 107.34, 107.78], u'A4.07': [80.81, 88.55, 80.69], u'A4.04': [96.97, 101.97, 90.51], u'A4.05': [101.6, 103.31, 93.43], u'A4.02': [106.75, 117.04, 116.23], u'A4.03': [89.07, 91.37, 88.49], u'A DYMOWS SZYMEL SZEMBE': [89.7, 96.9, 89.5], u'A4.01': [102.07, 72.37, 96.5], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [95.55, 96.22, 91.68], u'A4': [97.28, 98.7, 94.35]}}}
    UNITS_data = {u'roswera': {'SUVARDIO BRAND': {u'A4.06': [277.0, 303.0, 331.0], u'A4.07': [250.0, 225.0, 180.0], u'A4.04': [109.0, 97.0, 155.0], u'A4.05': [439.0, 399.0, 457.0], u'A4.02': [297.0, 230.0, 258.0], u'A4.03': [376.0, 376.0, 413.0], u'A DYMOWS SZYMEL SZEMBE': [11997.0, 10464.0, 12097.0], u'A4.01': [146.0, 124.0, 141.0], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [530.0, 443.0, 583.0], u'A4': [2424.0, 2197.0, 2518.0]}, 'ROMAZIC BRAND': {u'A4.06': [104.0, 238.0, 98.0], u'A4.07': [167.0, 170.0, 106.0], u'A4.04': [58.0, 119.0, 92.0], u'A4.05': [180.0, 157.0, 229.0], u'A4.02': [116.0, 123.0, 142.0], u'A4.03': [82.0, 112.0, 108.0], u'A DYMOWS SZYMEL SZEMBE': [6917.0, 7761.0, 7755.0], u'A4.01': [38.0, 18.0, 41.0], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [393.0, 401.0, 461.0], u'A4': [1138.0, 1338.0, 1277.0]}, 'ZAHRON BRAND': {u'A4.06': [200.0, 160.0, 192.0], u'A4.07': [285.0, 247.0, 303.0], u'A4.04': [171.0, 99.0, 158.0], u'A4.05': [278.0, 275.0, 322.0], u'A4.02': [136.0, 151.0, 221.0], u'A4.03': [265.0, 282.0, 326.0], u'A DYMOWS SZYMEL SZEMBE': [9542.0, 8657.0, 12991.0], u'A4.01': [58.0, 79.0, 100.0], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [399.0, 488.0, 584.0], u'A4': [1792.0, 1781.0, 2206.0]}, 'ZARANTA BRAND': {u'A4.06': [82.0, 76.0, 75.0], u'A4.07': [73.0, 85.0, 59.0], u'A4.04': [50.0, 43.0, 54.0], u'A4.05': [104.0, 74.0, 103.0], u'A4.02': [52.0, 63.0, 70.0], u'A4.03': [125.0, 120.0, 197.0], u'A DYMOWS SZYMEL SZEMBE': [6708.0, 5877.0, 6381.0], u'A4.01': [51.0, 24.0, 21.0], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [181.0, 147.0, 191.0], u'A4': [718.0, 632.0, 770.0]}, 'ROSWERA BRAND': {u'A4.06': [939.0, 731.0, 842.0], u'A4.07': [433.0, 378.0, 470.0], u'A4.04': [369.0, 300.0, 448.0], u'A4.05': [897.0, 748.0, 966.0], u'A4.02': [470.0, 551.0, 540.0], u'A4.03': [525.0, 466.0, 547.0], u'A DYMOWS SZYMEL SZEMBE': [27421.0, 23482.0, 26434.0], u'A4.01': [197.0, 199.0, 167.0], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [1292.0, 1069.0, 1299.0], u'A4': [5122.0, 4442.0, 5279.0]}}, u'atoris': {'ATORVASTEROL TOTAL': {u'A4.06': [282.0, 194.0, 205.0], u'A4.07': [212.0, 158.0, 227.0], u'A4.04': [115.0, 72.0, 115.0], u'A4.05': [307.0, 248.0, 273.0], u'A4.02': [117.0, 73.0, 89.0], u'A4.03': [152.0, 117.0, 172.0], u'A DYMOWS SZYMEL SZEMBE': [11546.0, 9925.0, 10559.0], u'A4.01': [61.0, 35.0, 34.0], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [536.0, 417.0, 357.0], u'A4': [1782.0, 1314.0, 1472.0]}, 'TORVACARD TOTAL': {u'A4.06': [None, None, None], u'A4.07': [None, None, None], u'A4.04': [None, None, None], u'A4.05': [None, None, None], u'A4.02': [None, None, None], u'A4.03': [None, None, None], u'A DYMOWS SZYMEL SZEMBE': [None, None, None], u'A4.01': [None, None, None], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [None, None, None], u'A4': [None, None, None]}, 'TORVALIPIN BRAND': {u'A4.06': [None, None, None], u'A4.07': [None, None, None], u'A4.04': [None, None, None], u'A4.05': [None, None, None], u'A4.02': [None, None, None], u'A4.03': [None, None, None], u'A DYMOWS SZYMEL SZEMBE': [None, None, None], u'A4.01': [None, None, None], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [None, None, None], u'A4': [None, None, None]}, 'ATORIS FILM C.TABS 30 MG 30': {u'A4.06': [29.0, 14.0, 20.0], u'A4.07': [28.0, 59.0, 6.0], u'A4.04': [18.0, 6.0, 25.0], u'A4.05': [18.0, 21.0, 16.0], u'A4.02': [4.0, 17.0, 5.0], u'A4.03': [15.0, 13.0, 13.0], u'A DYMOWS SZYMEL SZEMBE': [1291.0, 1067.0, 1230.0], u'A4.01': [7.0, 12.0, 13.0], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [52.0, 33.0, 43.0], u'A4': [171.0, 175.0, 141.0]}, 'ATORIS FILM C.TABS 30 MG 60': {u'A4.06': [2.0, 1.0, 2.0], u'A4.07': [10.0, 11.0, 10.0], u'A4.04': [4.0, 1.0, 7.0], u'A4.05': [4.0, 10.0, 11.0], u'A4.02': [2.0, 7.0, -7.0], u'A4.03': [3.0, 5.0, 4.0], u'A DYMOWS SZYMEL SZEMBE': [278.0, 325.0, 358.0], u'A4.01': [2.0, 0.0, 3.0], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [9.0, 14.0, 13.0], u'A4': [36.0, 49.0, 43.0]}, 'TULIP TOTAL': {u'A4.06': [352.0, 298.0, 387.0], u'A4.07': [431.0, 375.0, 396.0], u'A4.04': [160.0, 152.0, 198.0], u'A4.05': [518.0, 460.0, 548.0], u'A4.02': [270.0, 242.0, 254.0], u'A4.03': [260.0, 255.0, 263.0], u'A DYMOWS SZYMEL SZEMBE': [16285.0, 13699.0, 15166.0], u'A4.01': [118.0, 79.0, 84.0], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [622.0, 498.0, 577.0], u'A4': [2731.0, 2359.0, 2707.0]}, 'ATORIS TOTAL': {u'A4.06': [1001.0, 745.0, 903.0], u'A4.07': [726.0, 637.0, 741.0], u'A4.04': [465.0, 402.0, 428.0], u'A4.05': [1351.0, 1117.0, 1251.0], u'A4.02': [667.0, 631.0, 710.0], u'A4.03': [741.0, 626.0, 722.0], u'A DYMOWS SZYMEL SZEMBE': [34129.0, 29581.0, 31876.0], u'A4.01': [266.0, 178.0, 236.0], u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': [None, None, None], u'A4.08': [1649.0, 1362.0, 1399.0], u'A4': [6866.0, 5698.0, 6390.0]}}}
    CEGLY_data = {u'A4.06': {}, u'A4.07': {}, u'A4.04': {}, u'A4.05': {}, u'A4.02': {}, u'A4.03': {}, u'A DYMOWS SZYMEL SZEMBE': {}, u'A4.01': {}, u'GEOGRAPHY,PRODUCT BY MEASURES,PERIOD': {}, u'A4.08': {}, u'A4': {}}

    print "output1 START"
    output1 = Write_output_PnVa_Units_paternA(output_zakladki, PnVa_data)

    output1.start_output()
    output1.write_base_data()
    output1.write_PnVa_UNITS_data()
    output1.save_output()

    print "END output1"

    # output1.open_input_file()        # otwarcie pliku input z danymi
    # output1.get_date()               # odczytanie daty z pliku input z danymi
    # output1.get_PnVa_Units_column()  # znalezienie kolumn zawierających dane PnVa oraz UNITS
    # output1.get_cegla_positions()    # zbiera lokalizacje danych dla poszczególnych cegieł w zakładkach
    # output1.get_Cegly_data()
    # output1.get_PnVa_UNITS_data('ProfilTestowy2')

    # print 50 * "$$$"
    # print "output1.output_zakladki", output1.output_zakladki
    # print 50 * "$$$"

    # input2 = ReadInput(input_file_path2)
    # input2.open_input_file()        # otwarcie pliku input z danymi
    # input2.get_date()               # odczytanie daty z pliku input z danymi
    # input2.get_PnVa_Units_column()  # znalezienie kolumn zawierających dane PnVa oraz UNITS
    # input2.get_cegla_positions()    # zbiera lokalizacje danych dla poszczególnych cegieł w zakładkach
    # input2.get_Cegly_data()
    # input2.get_PnVa_UNITS_data('ProfilTestowy3')

    # print "output2 START"
    #
    # output2 = Write_output_CEGLY_paternA()
    # output2.start_output()
    # output2.write_base_data()
    # output2.write_PnVa_UNITS_data()
    # output2.save_output()
    # print "END output2"
