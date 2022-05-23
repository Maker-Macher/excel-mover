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

    dates = assign_dates(import_sheet)  # From import table

    table = create_table(export_sheet, name)
    # Allgemeine Tabelle wo jede Auftragsnummer steht und wer daran gearbeitet hat in der exportierten Tabelle

    auftr = get_auftrg_cells(import_sheet)  # Koordinaten jeder Auftragsnummer in der Tabelle

    copy_export_to_import(import_sheet, table, dates, auftr)


def copy_export_to_import(_import_sheet, _table, _dates, _coord):
    _coord_list = list(_coord.keys())
    _start, _end = None, None

    op_file = op.open(r'MöglicheProjekte_2022.xlsx', read_only=False, keep_vba=True)
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
                _mitarbeiter[_import_sheet[j][0]] = j
                # _mitarbeiter.append(_import_sheet[j][0])

            try:
                for k in _table[_coord_list[i]]:
                    _name = k[0]
                    _date = k[1]
                    _std = k[2]

                    _loc = [_dates[_date], _mitarbeiter[_name]]

                    # _import_sheet.at[_loc[0], _loc[1]] = _std
                    op_sheet.cell(row=_loc[0] + 1, column=_loc[1] + 1).value = _std

            except KeyError:
                continue

        except KeyError:
            continue

    op_file.save('export.xlsm')
    print("Done writing")


def get_auftrg_cells(_sheet):
    _dict = {}

    _df = _sheet.iloc[1]
    _nr_array = _df.to_numpy()

    c = 0
    for i in _nr_array:
        if not pd.isnull(i):
            try:
                i = int(i)
                _dict[i] = [c, 1]  # coord of Auftragsnummer, y always 0
            except ValueError as e:
                logging.debug(e)
        c += 1

    return _dict


def create_table(_sheet, _name_dict):
    _dict = {}

    for i in range(1, _sheet.shape[0]):
        _val = _sheet.iloc[i]

        if pd.isnull(_val["EmpfAuftrg"]):
            continue

        _auftr = int(_val["EmpfAuftrg"])
        # _name = _val["Name"]
        _name = _name_dict[_val["Name"]]
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
    for i in range(20, 276):
        _val = _sheet[0][i]

        if not pd.isnull(_val):
            _dict[_val] = i

    return _dict


if __name__ == '__main__':
    main()
