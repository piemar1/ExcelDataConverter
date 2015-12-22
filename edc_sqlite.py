# ! /usr/bin/env python
# -*- coding: utf-8 -*-
__author__ = 'Marcin Pieczyński'


import sqlite3 as lite


class SQliteEdit:
    def __init__(self):

        self.profiles_name_list = []
        self.con = lite.connect('Profile_database.db')  # Owieranie pliku bazy danych1
        self.cursor = self.con.cursor()

        self.output_zakladki = []
        self.output_leki = []
        self.output_leki_cegly = []
        self.output_lista_cegiel = []

    def get_tables_from_db(self):
        """ Zczytuje nazwy tebelek z pliku db SQlite potrzebne do wyboru profilu"""

        self.cursor.execute("SELECT name FROM sqlite_master WHERE type = 'table';")
        self.profiles_name_list = []

        lista = (self.cursor.fetchall())      # Zwraca listę krotek zawierającą stringi z nazwami tabel
        for elem in lista:
            self.profiles_name_list.append(elem[0])

        self.profiles_name_list = tuple(self.profiles_name_list)
        print self.profiles_name_list

    def get_data_from_profile(self, profil_to_use):
        """ Zczytuje dane dla wybranego profilu z pliku SQlite db
        potrzebne to filtorowania i konwersji danych"""

        print profil_to_use

        self.cursor.execute("SELECT * FROM " + profil_to_use)
        from_db = self.cursor.fetchall()

        for row in from_db:
            if row[0] == "PvNa_UNITS":
                self.output_zakladki.append(row[1])
                a = []
                b = row[2].split(",")
                for elem in b:
                    el = elem.strip()
                    a.append(el)
                self.output_leki.append(a)

            if row[0] == "CEGLY":
                b = row[2].split(",")
                for elem in b:
                    el = elem.strip()
                    self.output_leki_cegly.append(el)

            if row[0] == "lista_cegiel":
                nazwa_cegla_list = row[2]
                lista = nazwa_cegla_list.split(",")
                for elem in lista:
                    el = elem.strip()
                    self.output_lista_cegiel.append(el)

    def SQsave_profile(self, nazwa, final_profile):
        """Zapisuje profil o nazwie nazwa oraz zawartości final_profil w bazie SQlite"""

        self.cursor.execute("DROP TABLE IF EXISTS " + nazwa)
        self.cursor.execute("CREATE TABLE " + nazwa + "(Typ TEXT, Id_Grupy TEXT, lek TEXT)")
        self.cursor.executemany("INSERT INTO " + nazwa + " VALUES(?, ?, ?)", final_profile)

    def SQedit_profile(self,profil_to_edit):
        """ Odczytuje zawartość profilu """     # Metoda czeka na napisanie
        pass

    def SQdel_profile(self,profile_to_del):
        """ Usuwa wskazany profil profile_to_del z bazy SQlite"""
        self.cursor.execute("DROP TABLE IF EXISTS " + profile_to_del)


if __name__ == '__main__':
    m = SQliteEdit()
    m.get_tables_from_db()

    m.get_data_from_profile('ProfilTestowy1')

    print "output_zakladki  ", m.output_zakladki
    print "output_lista_cegiel  ", m.output_lista_cegiel

    for elem in m.output_leki:
        print "output_leki  ", elem

    print "output_leki_cegly    ", m.output_leki_cegly




    # print 50 * "XXX"
    # SQliteEdit().get_data_from_profile('ProfilTestowy2')
    # print 50 * "XXX"
    # SQliteEdit().get_data_from_profile('ProfilTestowy3')



