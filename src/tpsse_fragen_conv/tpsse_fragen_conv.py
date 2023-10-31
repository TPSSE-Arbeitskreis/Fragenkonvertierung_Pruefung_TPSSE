from pathlib import Path
from datetime import date
from datetime import datetime
import click
import sys
import yaml
from typing import Any, IO
from rich import print
import os
import json
# from .myloadenv import load_env
from openpyxl import load_workbook
from openpyxl.styles import Alignment


class FragenConv:
    
    def __init__(self, main_input, template, outfile):
        self._main_input = Path(main_input)
        self._today = date.today().strftime("%d.%m.%Y")
        with open(self._main_input, 'r', encoding="utf8") as f:
            self._data = yaml.load(f, Loader=yaml.CLoader)
        self._path = self._main_input.parent
        self._cwd = Path.cwd()

        self._outfile = Path(outfile)
        self._template = Path(template)
        self._load_fragen()

    @property
    def data(self):
        return self._data


    def _load_fragen(self):
        for le in self._data['lerneinheiten']:
            fragen_datei = self._path / le['datei']
            with open(fragen_datei, 'r', encoding="utf8") as f:
                temp = yaml.load(f, Loader=yaml.CLoader)
                le['fragenpool'] = temp


    def _copy_ws(self, wb, le):
        w = wb["Template SC_MC"]
        t = wb.copy_worksheet(w)
        t.title = f'{le["id"]:02d} SC_MC'

    def _write_ws(self, wb, le):
        if not le["fragenpool"]["fragen"]:
            return
        a = Alignment(wrap_text = True, shrinkToFit=True)
        w = wb[f'{le["id"]:02d} SC_MC']
        r = start_row = 2
        c = start_col = 2
        for f in le["fragenpool"]["fragen"]:
            w.cell(row=r, column=c).value = self._today
            w.cell(row=r, column=c+1).value = ("MC")
            w.cell(row=r, column=c+3).value = ("1")
            w.cell(row=r, column=c+4).value = f["text"]
            w.cell(row=r, column=c+4).alignment = a
            oc = c+5
            for o in f["optionen"]:
                w.cell(row=r, column=oc).value = o["text"]
                w.cell(row=r, column=oc).alignment = a
                if 'y' == o["richtig"]:
                    w.cell(row=r, column=oc+1).value = "x"
                oc += 2
            r += 1

    def _write_blueprint(self, wb):
        bp = wb["Blueprint"]
        r = start_row = 2
        c = start_col = 2
        for le in self._data["lerneinheiten"]:
            bp.cell(row=r, column=c).value = le["fragenpool"]["thema"]
            bp.cell(row=r, column=c+1).value = "MC"
            bp.cell(row=r, column=c+4).value = bp.cell(row=r, column=c+2).value = le["fragenpool"]["pruefungsfragen"]
            r += 1
            self._copy_ws(wb, le)
            self._write_ws(wb, le)
        wb["Template SC_MC"].sheet_state = 'hidden'


    def write_xls(self):
        wb = load_workbook(self._template)
        self._write_blueprint(wb)
        wb.save(self._outfile)


@click.command()
@click.option('--fmain', envvar='FRAGENCONV_MAIN', help="Main-Datei für Prüfungsfragen (FRAGENCONV_MAIN)")
@click.option('--template', envvar='FRAGENCONV_TEMPLATE', help="Main-Datei für Prüfungsfragen (FRAGENCONV_TEMPLATE)")
@click.option('--out', envvar='FRAGENCONV_OUT', help="Main-Datei für Prüfungsfragen (FRAGENCONV_OUT)")
def run(fmain, template, out):
    converter = FragenConv(fmain, template, out)
    converter.write_xls()

def main():
    # load_env()
    run()


if __name__ == '__main__':
    main()

