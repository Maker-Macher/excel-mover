"""
Author: Viktor Engelowski
Description: Read cells and rows from one Spreadsheet to insert the data into another
"""
import logging

import pandas as pd
import openpyxl as op
import names  # Lokale Datei mit Mitarbeiter Namen und deren "Kurzform"


def main():
    export_sheet = pd.read_excel(r'Export_ZeitenSAP_18052022.xlsx')
    import_sheet = pd.read_excel(r'MöglicheProjekte_2022.xlsx', sheet_name='Quartal 1', header=None)

    name = names.get_names()

    dates = assign_dates(import_sheet)

    table = create_table(export_sheet, name)
    # Allgemeine Tabelle wo jede Auftragsnummer steht und wer daran gearbeitet hat in der exportierten Tabelle

    auftr = get_auftrg_cells(import_sheet)  # Koordinaten jeder Auftragsnummer in der Tabelle

    copy_export_to_import(import_sheet, table, dates, auftr)


def copy_export_to_import(_import_sheet: pd.DataFrame, _table: dict, _dates: dict, _coord: dict) -> None:
    _coord_list = list(_coord.keys())  # Verwandelt Dict in ein Array
    _start, _end = None, None

    op_file = op.open(r'MöglicheProjekte_2022.xlsx', read_only=False, keep_vba=True)
    # Öffnet Datei mit einem anderen Package, um die Daten einzutragen
    op_sheet = op_file['Quartal 1']

    for i in range(len(_coord_list)):
        try:
            _start = _coord[_coord_list[i]][0]
            _end = _coord[_coord_list[i + 1]][0] - 1
        except IndexError:
            _end = _start + 100

        _mitarbeiter = {}  # Koordinaten an dem jeder Mitarbeiter eines Auftrages steht

        try:
            for j in range(_start + 2, _end + 1):
                _mitarbeiter[_import_sheet[j][0]] = j  # Ändert den "langen" Namen in den kurzen um

            try:
                for k in _table[_coord_list[i]]:
                    _name = k[0]
                    _date = k[1]
                    _std = k[2]

                    _loc = [_dates[_date], _mitarbeiter[_name]]
                    # Koordinaten des Datums zusammen mit dem Mitarbeiter ergeben die Zelle

                    op_sheet.cell(row=_loc[0] + 1, column=_loc[1] + 1).value = _std  # Schreibt Stunden in Zelle

            except KeyError:
                # Falls es den Auftrag in der Zieltabelle nicht gibt, fortsetzen
                continue

        except KeyError:
            # Wenn _end größer ist als die Anzahl der Mitarbeiter
            continue

    op_file.save('export.xlsm')
    print("Done writing new file")


def get_auftrg_cells(_sheet: pd.DataFrame) -> dict:
    _dict = {}

    _df = _sheet.iloc[1]
    _nr_array = _df.to_numpy()  # Nimmt die Zweite Reihe der Tabelle, wo alle Auftr. stehen und konvertiert zum Array

    c = 0
    for i in _nr_array:
        if not pd.isnull(i):
            try:
                i = int(i)  # Auftragsnummer als int abspeichern
                _dict[i] = [c, 1]  # Koordinaten der Auftragsnummer, y immer 1
            except ValueError:
                pass  # Wenn es keine Auftragsnummer ist, weiter machen
        c += 1

    return _dict


def create_table(_sheet: pd.DataFrame, _name_dict: dict) -> dict:
    _dict = {}

    for i in range(1, _sheet.shape[0]):  # shape[0] = die Anzahl an Reihen
        _val = _sheet.iloc[i]

        if pd.isnull(_val["EmpfAuftrg"]):
            continue  # Falls unter "EmpfAuftrg" nicht steht, zum nächsten

        _auftr = int(_val["EmpfAuftrg"])
        _name = _name_dict[_val["Name"]]  # Name wird als kurzform abgespeichert
        _date = _val["Datum"]
        _hrs = _val["Stunden"]

        if _auftr not in _dict.keys():
            _dict[_auftr] = [[_name, _date, _hrs]]  # Erstelle neuen Key, wenn Auftrag noch nicht im dict steht
        else:
            _dict[_auftr] = _dict[_auftr] + [[_name, _date, _hrs]]  # Ansonsten zum bestehenden Key anfügen

    return _dict


def assign_dates(_sheet: pd.DataFrame) -> dict:
    # Jedes Datum bekommt eine y-Koordinate zugewiesen
    _dict = {}
    # TODO: Automatisch dir range erkennen
    for i in range(20, 276):  # Range der Daten in der Zieltabelle
        _val = _sheet[0][i]

        if not pd.isnull(_val):
            _dict[_val] = i  # Wenn die Zelle nicht leer ist, in ein dict stecken

    return _dict


if __name__ == '__main__':
    main()
