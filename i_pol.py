# import pandas as pd
import xlrd
from rdflib import Graph, Namespace, FOAF, XSD, RDF, RDFS
import os.path
from enum import Enum

SUBDIR = "./iec/pollutions/"

FILES = { 'Данные бар Ярки Сев Байкал.xls':1,
          'Данные Харанцы Ольхон.xls':2,
          'Данные Хужир Ольхон.xls':3 }

EQUIPMENT = "S8 Tiger"
EQUIPMENT_TYPE = "X-Ray fluorescence analysis"

PT = Namespace('https://crust.irk.ru/ontology/pollution/terms/1.0/')
P = Namespace('https://crust.irk.ru/ontology/pollution/1.0/')

G = Graph(bind_namespaces="rdflib")
G.bind('pt', PT)
G.bind('pi', P)

class State(Enum):
    NONE = 0
    CLASS = 1
    HEADER = 2
    DATA = 3
    IGNORE = 4

class ImpState:
    def __init__(self):
        self.state = State.NONE
        self.instr = False

    def data(self):
        return self.state == State.DATA

    def r(self, row, col):
        f = row[0]
        fs = f.value

        if f.ctype == xlrd.XL_CELL_TEXT:
            fss = fs.strip()
            if fss.startswith("#"):
                fss = fss.lstrip("#")
                try:
                    self.state = State[fss]
                except KeyError as k:
                    print("#! ERROR Key: {}".format(k))
                    quit()
                self.instr = True

        if self.data() and not self.instr:
            # print("DT: {}".format(row))
            for i, cell in enumerate(row):
                self.c(cell, i, col)

        self.instr = False

    def c(self, cell, row, col):
        pass


def parse_sheet(sh):
    # print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
    st = ImpState()
    for rx in range(sh.nrows):
        st.r(sh.row(rx), rx)


def parse_xl(file):
    print("# FILE: {} at {}".format(file, SUBDIR))
    pathfile = os.path.join(SUBDIR, file)
    # df = pd.read_excel(pathfile)
    wb = xlrd.open_workbook(pathfile)
    print("# Sheet names: {}".format(wb.sheet_names()))
    for sheet in wb.sheet_names():
        print("# Wb: {}, sheet: {}". format(file, sheet))
        sh = wb.sheet_by_name(sheet)
        print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
        parse_sheet(sh)


if __name__=="__main__":
    for file in FILES:
        parse_xl(file)
