# import pandas as pd
import xlrd
from rdflib import Graph, Namespace, FOAF, XSD, RDF, RDFS, URIRef, Literal
import os.path
import unicodedata
from enum import Enum

ONTODIR = "./iec/"
SUBDIR = "./iec/pollutions/"


TARGET = 'import.ttl'
TARGETMT = 'PeriodicTable.ttl'

EQUIPMENT = "S8 Tiger"
EQUIPMENT_TYPE = "X-Ray fluorescence analysis"

PT = Namespace('https://crust.irk.ru/ontology/pollution/terms/1.0/')
P = Namespace('https://crust.irk.ru/ontology/pollution/1.0/')
MT = Namespace('http://www.daml.org/2003/01/periodictable/PeriodicTable#')

G = Graph(bind_namespaces="rdflib")
GMT = Graph(bind_namespaces="rdflib")
GMT.parse(location=os.path.join(ONTODIR, 'PeriodicTable.owl'))
GS = [G, GMT]

for _ in GS:
    _.bind('pt', PT)
    _.bind('pi', P)
    _.bind('mt', MT)

def normURI(s):
    r = ""
    for c in s:
        d = unicodedata.category(c)
        if d[0] in ['L','N']:
            r+=c
        elif len(r)>0 and r[-1]!='_':
            r+='_'
    return r.rstrip('_')

class State(Enum):
    NONE = 0
    CLASS = 1
    HEADER = 2
    DATA = 3
    IGNORE = 4

class ImpState:
    _sample_iri_ = PT['sample']
    def __init__(self, graph, datasetIRI):
        self.state = State.NONE
        self.instr = False
        self.cls = {}
        self.header = {}
        self.dsiri = datasetIRI
        self.g = graph
        self.sample = None
        self.sample_col = None

    def data(self):
        return self.state == State.DATA
    def ign(self):
        return self.state == State.IGNORE

    def r(self, row, rx):
        self.instr = False
        self.sample = None

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
        if self.ign():
            return

        if self.data() and not self.instr:
            # print("DT: {}".format(row))
            if self.sample is None:
                # print(self.header)
                self.c(row[self.sample_col], rx, self.sample_col)
            for i, cell in enumerate(row):
                self.c(cell, rx, i)
            return

        if self.state == State.HEADER and not self.instr:
            for i, cell in enumerate(row):
                self.h(cell, rx, i)
            return


    def c(self, cell, row, col):
        if cell.ctype in [xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK]:
            return
        try:
            field = self.header[col]
        except KeyError as k:
            print("#! ERROR header key {} not in {}".format(k, self.header))
            # quit()
            return
        add = self.g.add
        ds = self.dsiri
        if field == self._sample_iri_:
            self.sample = P[cell.value.replace(' ','')]
            add((ds, self._sample_iri_, self.sample))
        elif self.sample is not None:
            add((self.sample, PT[field], Literal(cell.value)))
        else:
            print('#! ERROR: nowhere to store {} R:{} C:{}\n#!{}'.format(cell, row, col, self.header))
            quit()

    def h(self, cell, row, col):
        # print(cell, col)
        if cell.ctype in [xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK]:
            return

        name = normURI(str(cell.value))
        add = self.g.add
        ds = self.dsiri
        if name in self._sample_names_:
            name = self._sample_iri_
            self.sample_col = col

        self.header[col] = name

class Yarki(ImpState):
    _sample_names_ = ['Sample', 'Полевой']

class Kharantsy(ImpState):
    _sample_names_ = ['Номер_пробы']

class Khuzhir(Kharantsy):
    pass


FILES = { 'Данные бар Ярки Сев Байкал.xls': Yarki,
          'Данные Харанцы Ольхон.xls': Kharantsy,
          'Данные Хужир Ольхон.xls': Khuzhir }


def parse_sheet(sh, sheetIRI, constr):
    # print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
    st = constr(G, sheetIRI)
    for rx in range(sh.nrows):
        st.r(sh.row(rx), rx)


def parse_xl(file, constr):
    print("# FILE: {} at {}".format(file, SUBDIR))
    pathfile = os.path.join(SUBDIR, file)
    # df = pd.read_excel(pathfile)
    wb = xlrd.open_workbook(pathfile)
    print("# Sheet names: {}".format(wb.sheet_names()))
    for sheet in wb.sheet_names():
        print("# Wb: {}, sheet: {}". format(file, sheet))
        sh = wb.sheet_by_name(sheet)
        sheetname = normURI(file+"_"+sheet)
        print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
        parse_sheet(sh, P[sheetname], constr)


if __name__=="__main__":
    for file, constr in FILES.items():
        parse_xl(file, constr)
    with open(TARGET, "w") as o:
        o.write(G.serialize(format='turtle'))
    with open(TARGETMT, "w") as o:
        o.write(GMT.serialize(format='turtle'))
