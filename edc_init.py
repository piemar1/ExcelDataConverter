# ! /usr/bin/env python
# -*- coding: utf-8 -*-
__author__ = 'Marcin Pieczyński'


"""
Moduł służący do inicjaizacji programu ExcelDataConverter.
Sprawdza datę ważnoścli licencji oraz czy program może być uruchomiony na danym komputerze.
Ostatecznie uruchamia GUI programu z pliku EDC_MainGui.py.
"""

import uuid
import datetime
import cPickle as pickle
import tkMessageBox
import edc_main_gui

# print socket.gethostname()


class LicenceCheck:
    """Klasa zawierające metody sprawdzające czy na PC można uruchomić program:
    - sprawdza czy licencja obejmuje dany PC
    - sprawdza czy licencja nie wygasła
    """
    def __init__(self):
        self.expiry_date = [2015, 12, 28]    # WAŻNE data ważności licencji na program
        self.my_pc_win = [132098963232679]   # id mojego kompa  --> 132098963232679  NIE POTRZEBNE
        self.my_pc_lin = [272774785408889]
        self.pc_license_list = None          # lista pc z licencją
        self.delta = None

        self.licence_is_valid = False
        self.license_for_this_pc = False     # LICENCJA NA TEN KOMPUTER - Domyślnie False

        self.check_ip()
        self.count_exipy_days()
        this_pc = uuid.getnode()

    def check_ip(self):
        """ Metoda sprawdzająca licencję dla tego PC."""

        """ TU GDZIEŚ JEST BŁĄD, zablokowało mój własny komputer po uruchomieniu skryptu na linuksie!!!"""
        """ wcześniej na windowskie chodziło !!!"""
        try:
            self.pc_license_list = pickle.load(open("license.p", "rb"))
            # print self.pc_license_list

            this_pc = uuid.getnode()

            # Sprawdzenie czy jest licencja na pc
            if len(self.pc_license_list) == 1:
                if this_pc in self.pc_license_list:
                    pass
                if this_pc not in self.pc_license_list:
                    self.pc_license_list.append(this_pc)
                self.license_for_this_pc = True              # Licencja jest ok

            elif len(self.pc_license_list) == 2:
                if this_pc in self.pc_license_list:
                    self.license_for_this_pc = True          # Licencja jest ok
                else:
                    self.license_for_this_pc = False         # Licencja jest NIEWAŻNA
        except IOError:
            pass
        print self.pc_license_list

    def count_exipy_days(self):
        """ Oblicza liczbę dni do zakończenia ważności licencji """

        ex_year, ex_month, ex_day = self.expiry_date
        now = datetime.datetime.now()
        current_year, current_month, current_day = now.year, now.month, now.day

        d0 = datetime.date(ex_year, ex_month, ex_day)
        d1 = datetime.date(current_year, current_month, current_day)
        self.delta = (d0 - d1).days
        # print "Liczba dni ważnej licencji - ", self.delta, "dni."
        if self.delta >= 0:
            self.licence_is_valid = True



if __name__ == "__main__":

    license = LicenceCheck()

    if license.license_for_this_pc and license.licence_is_valid:
        # URUCHOMIENIE PROGRAMU
        from edc_main_gui import MainGui
        MainGui()

    elif license.license_for_this_pc is False:
        # print "Brak ważne licencji na ten komputer."
        tkMessageBox.showerror("Brak ważnej licencji.",
                               "Nie posiadasz licencji na uruchomiwnie programu Excel Data Converter."
                               "W celu korzystania z programu skontaktuj się z Marcinem Pieczyńskim."
                               "\n\n marcin-pieczynski@wp.pl")
    elif license.licence_is_valid is False:
        # print "Licencja na użytkowanie programu skończyła się."
        tkMessageBox.showerror("Brak ważnej licencji",
                               "Licencja na użytkowanie programu Excel Data Converter zakończyła się. "
                               "W celu dalszego korzystania z programu skontaktuj się z Marcinem Pieczyńskim."
                               "\n\n marcin-pieczynski@wp.pl")


