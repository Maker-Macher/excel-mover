"""
Author: Viktor Engelowski
Description: Read cells and rows from one Spreadsheet to insert the data into another
"""

import pandas as pd
import numpy as np


def main():
    export_sheet = pd.read_excel(r'Export.xlsx')
    import_sheet = pd.read_excel(r'Mappe.xlsx')  # sheet_name='..'

    name = {"Hans Peter": "Hans",
            "Frank": "Frank",
            "Stephan": "Stephan",
            "Markus von Lichtenstein": "Markus"}

    dates = assign_dates(import_sheet)  # From import table
    # print(dates)

    table = create_table(export_sheet)
    # Allgemeine Tabelle wo jede Auftragsnummer steht und wer daran gearbeitet hat in der exportierten Tabelle
    # print(table)

    auftr = get_auftrg_cells(import_sheet)  # Koordinaten jeder Auftragsnumer in der Tabelle
    # print(auftr)

    copy_export_to_import(import_sheet, table, dates, auftr)


def copy_export_to_import(_import_sheet, _table, _dates, _coord):
    _coord_list = list(_coord.keys())
    _start, _end = None, None

    for i in range(len(_coord_list)):
        try:
            _start = _coord[_coord_list[i]][0]
            _end = _coord[_coord_list[i + 1]][0] - 1
        except IndexError:
            _end = _start + 100

        # TODO: Finish finding the cells where the names and dates are from _table


def get_auftrg_cells(_sheet):
    _dict = {}

    _df = _sheet.iloc[0]
    _nr_array = _df.to_numpy()

    for i in _nr_array:
        if not pd.isnull(i):
            i = int(i)
            _dict[i] = [np.where(_nr_array == i)[0][0], 0]  # coord of Auftragsnummer, y always 0

    return _dict


def create_table(_sheet):
    _dict = {}
    for i in range(_sheet.shape[0]):
        _val = _sheet.iloc[i]

        if pd.isnull(_val["EmpfAuftrg"]):
            continue

        _auftr = int(_val["EmpfAuftrg"])
        _name = _val["Name"]
        _date = _val["Datum"]
        _hrs = _val["Stunden"]

        if _auftr not in _dict.keys():
            _dict[_auftr] = [[_name, _date, _hrs]]
        else:
            _dict[_auftr] = _dict[_auftr] + [[_name, _date, _hrs]]

    return _dict


def assign_dates(_sheet):
    # Assign each date an index
    _dict = {}
    for i in range(5, _sheet.shape[0]):
        val = _sheet["Unnamed: 0"][i]

        _dict[val] = i

    return _dict


if __name__ == '__main__':
    main()
