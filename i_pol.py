# import pandas as pd
import xlrd
from rdflib import Graph, Namespace, FOAF, XSD, RDF, RDFS, URIRef, Literal, BNode
import os.path
import unicodedata
from enum import Enum
import re

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

ELRE = re.compile(r'^([A-Z][a-z]{,3})(.*)$')

for _ in GS:
    _.bind('pt', PT)
    _.bind('pi', P)
    _.bind('mt', MT)

ElToIRI = {}
GeoSample = PT.GeologicalSample
GeoMeasure = PT.Measurement
samplerel = PT.sample
PPM = PT.PPM
Percent = PT.Percent

G.add((PT.PPM, RDFS.label, Literal('Грам/Моль', lang='ru')))
G.add((PT.Percent, RDFS.label, Literal('Процент', lang='ru')))

for el in GMT.subjects(RDF.type, MT.Element):
    ElToIRI[el.fragment] = MT[el]

# print(ElToIRI)

def normURI(s):
    r = ""
    for c in s:
        d = unicodedata.category(c)
        if d[0] in ['L','N']:
            r+=c
        elif len(r)>0 and r[-1]!='_':
            r+='_'
    return r.rstrip('_')


def elem(name):
    # import pudb; pu.db
    return ElToIRI.get(name, None)


class State(Enum):
    NONE = 0
    CLASS = 1
    HEADER = 2
    DATA = 3
    IGNORE = 4
    LOCATION = 5
    DETLIM = 6


class ImpState:
    _sample_iri_ = samplerel
    _belongs_iri_ = PT['location']

    def __init__(self, graph, datasetIRI, locations=None):
        self.state = State.NONE
        self.instr = False
        self.cls = {}
        self.header = {}
        self.dsiri = datasetIRI
        self.g = graph
        self.sample = None
        self.locations = []
        self.sample_col = None
        self.proc_locs(locations)
        self.loc_fence = len(self.locations)

    def proc_comp(self, name, value):
        name = name.strip()
        mo = ELRE.match(name)
        add = self.g.add
        rel = PT[normURI(name)]
        if mo is None:
            add((self.sample, rel, Literal(value)))
            return
        el = mo.group(1)
        eliri = elem(el)
        if eliri is None:
            add((self.sample, rel, Literal(value)))
            return
        # print(eliri)
        rest = mo.group(2)
        el = elem(el)
        m = BNode()
        add((self.sample, PT.measure, m))
        add((m, RDF.type, GeoMeasure))
        add((m, MT.element, el))
        add((m, PT.value, Literal(value)))
        add((m, PT.unit, PPM))

    def proc_locs(self, locations):
        if locations is None:
            return
        for loc in locations:
            self.proc_loc(loc)

    def proc_loc(self, location):
        location = location.strip()
        uriloc = P[normURI(location)]
        locs = self.locations
        if locs:
            prev = locs[-1]
            self.belongs(uriloc, prev)
        add = self.g.add
        add((uriloc, RDFS.label, Literal(location, lang='ru')))
        self.locations.append(uriloc)

    def belongs(self, obj, location=None):
        if location is None:
            if len(self.locations) > 0:
                location = self.locations[-1]
            else:
                return  # None to belong to
        add = self.g.add
        # print(self.locations)
        add((obj, self._belongs_iri_, location))

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
        if self.ign() or self.instr:
            return

        if self.data():
            # print("DT: {}".format(row))
            if self.sample is None:
                # print(self.header)
                self.c(row[self.sample_col], rx, self.sample_col)
            for i, cell in enumerate(row):
                self.c(cell, rx, i)
            return

        if self.state == State.HEADER:
            for i, cell in enumerate(row):
                self.h(cell, rx, i)
            return

        if self.state == State.LOCATION:
            for cell in row:
                if cell.ctype == xlrd.XL_CELL_TEXT:
                    if len(self.locations) > self.loc_fence:
                        self.locations.pop()
                    loc = cell.value.strip()
                    self.proc_loc(loc)

        # TODO: DETLIM

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
            add((self.sample, RDF.type, GeoSample))
            self.belongs(self.sample)
        elif self.sample is not None:
            self.proc_comp(field, cell.value)
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


FILES = {'Данные бар Ярки Сев Байкал.xls': (Yarki, ['Северный Байкл', 'Ярки']),
         'Сводная таблица РФА_Бураевская площадь № 70-2024.xls': (Yarki, ['Бураевская площадь № 70']),
         'Данные Харанцы Ольхон.xls': (Kharantsy, ['Ольхон', 'Харанцы']),
         'Данные Хужир Ольхон.xls': (Khuzhir, ['Ольхон', 'Хужир'])}


def parse_sheet(sh, sheetIRI, comp):
    # print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
    constr, locs = comp
    st = constr(G, sheetIRI, locations=locs)
    for rx in range(sh.nrows):
        st.r(sh.row(rx), rx)


def parse_xl(file, comp):
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
        parse_sheet(sh, P[sheetname], comp)


if __name__=="__main__":
    for file, comp in FILES.items():
        parse_xl(file, comp)
    with open(TARGET, "w") as o:
        o.write(G.serialize(format='turtle'))
    with open(TARGETMT, "w") as o:
        o.write(GMT.serialize(format='turtle'))
