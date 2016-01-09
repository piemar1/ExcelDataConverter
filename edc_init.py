# ! /usr/bin/env python
# -*- coding: utf-8 -*-
import uuid
import datetime
import cPickle as pickle
import tkMessageBox


__author__ = 'Marcin Pieczyński'


"""
Module for initialization Excel Data Converter application.
After launching it checks if licence for this pc is valid.
Eventually it launches GUI from edc_main_gui.py file.
"""


class LicenceCheck:
    """
    The Class containing methods for checking if the licence is valid for this pc and for the date of use.
    """
    def __init__(self):
        self.expiry_date = [2016, 1, 20]    # IMPORTANT  - Expiration of license
        self.my_pc_win = [132098963232679]   # id of my PC-win
        self.my_pc_lin = [272774785408889]   # id of my PC-lin
        self.pc_license_list = None          # list with pc with valid licence
        self.delta = None

        self.licence_is_valid = False
        self.license_for_this_pc = False

        self.check_ip()
        self.count_exipy_days()

    def check_ip(self):
        """
        The method for checking licence for this pc.
        """
        try:
            self.pc_license_list = pickle.load(open("license.p", "rb"))
            this_pc = uuid.getnode()

            # checking licence for this pc
            if len(self.pc_license_list) <= 1:
                if this_pc in self.pc_license_list:
                    pass
                if this_pc not in self.pc_license_list:
                    self.pc_license_list.append(this_pc)
                self.license_for_this_pc = True              # Licence is OK

            elif len(self.pc_license_list) == 2:
                if this_pc in self.pc_license_list:
                    self.license_for_this_pc = True          # Licence is OK
                else:
                    self.license_for_this_pc = False         # Licence is not valid
        except IOError:
            pass

    def count_exipy_days(self):
        """
        The method for counting the number of days of validity of license.
        """
        ex_year, ex_month, ex_day = self.expiry_date
        now = datetime.datetime.now()
        current_year, current_month, current_day = now.year, now.month, now.day

        d0 = datetime.date(ex_year, ex_month, ex_day)
        d1 = datetime.date(current_year, current_month, current_day)

        self.delta = (d0 - d1).days                             # number of days of validity of license
        if self.delta >= 0:
            self.licence_is_valid = True

    @staticmethod
    def no_licence():
        """
        The window with message informing of lacking of licence for this PC.
        """
        tkMessageBox.showerror("Brak ważnej licencji.",
                               "Nie posiadasz licencji na uruchomiwnie programu Excel Data Converter. "
                               "W celu korzystania z programu skontaktuj się z Marcinem Pieczyńskim."
                               "\n\n marcin-pieczynski@wp.pl")

    @staticmethod
    def licence_no_valid():
        """
        The window with message informing that licence has expired.
        """
        tkMessageBox.showerror("Brak ważnej licencji",
                               "Licencja na użytkowanie programu Excel Data Converter zakończyła się. "
                               "W celu dalszego korzystania z programu skontaktuj się z Marcinem Pieczyńskim."
                               "\n\n marcin-pieczynski@wp.pl")

if __name__ == "__main__":

    lic = LicenceCheck()

    if lic.license_for_this_pc and lic.licence_is_valid:
        from edc_main_gui import MainGui
        MainGui()

    elif lic.license_for_this_pc is False:
        lic.no_licence()

    elif lic.licence_is_valid is False:
        lic.licence_no_valid()
