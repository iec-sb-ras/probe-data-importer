# import pandas as pd
import xlrd
from rdflib import Graph, Namespace, FOAF, XSD, RDF, RDFS, URIRef, Literal, BNode
from rdflib.namespace import WGS, SDO
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

PT = Namespace('http://crust.irk.ru/ontology/pollution/terms/1.0/')
P = Namespace('http://crust.irk.ru/ontology/pollution/1.0/')
MT = Namespace('http://www.daml.org/2003/01/periodictable/PeriodicTable#')
# GEO = Namespace('http://www.w3.org/2003/01/geo/wgs84_pos#')

G = Graph(bind_namespaces="rdflib")
GMT = Graph(bind_namespaces="rdflib")
GMT.parse(location=os.path.join(ONTODIR, 'PeriodicTable.owl'))
GS = [G, GMT]

COMPRE = re.compile(r'^(([A-Z][a-z]{,2}\d{,2})+)(.*?)$')
COMPELRE = re.compile(r'[A-Z][a-z]{,2}\d{,2}')
# r'([A-Z][a-z]*)(\d+(?:\.\d+)?)?'
ELRE = re.compile(r'^([A-Z][a-z]{,2})+(.*)$')

DEGRE = re.compile(r'^(\d{,3})(\D)(\d{,2})(\D)(\d{,2}(\.\d{,2})?)(\D)(.*)$')
# 107°25'28.48"В
# m = DEGRE.match('107°25\'28.48"В')
# print(m, m.groups())
# m = DEGRE.match('107°25\'28"В')
# print(m, m.groups())
# quit()

# spl = COMPELRE.findall("C2H5W8K9OH")
# print(spl)
# spl = COMPELRE.findall("OH")
# print(spl)
# spl = COMPELRE.findall("SiO2")
# print(spl)
# # quit()

COMPOUNDS = {}

for _ in GS:
    _.bind('pt', PT)
    _.bind('pi', P)
    _.bind('mt', MT)
    # _.bind('geo', GEO)

ElToIRI = {}
GeoSample = PT.GeologicalSample
GeoMeasure = PT.Measurement
samplerel = PT.sample
PPM = PT.PPM
Percent = PT.Percent
SpatialThing = WGS.SpatialThing
Long = WGS.long
Lat = WGS.lat
IgnitionLosses = PT.IgnitionLosses

G.add((PT.PPM, RDFS.label, Literal('мг/кг', lang='ru')))
G.add((PT.Percent, RDFS.label, Literal('Процент', lang='ru')))

for el in GMT.subjects(RDF.type, MT.Element):
    ElToIRI[el.fragment] = el

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
        self.prev_class = None
        self.cls = {}
        self.header = {}
        self.dsiri = datasetIRI
        self.g = graph
        self.sample = None
        self.locations = []
        self.sample_col = None
        self.proc_locs(locations)
        self.loc_fence = len(self.locations)
        self.dlims = {}

    def proc_comp(self, names, value, delim=False):
        uvalue = str(value).upper().strip()
        if isinstance(value, str):
            value = value.strip()
        if 'N.A.' in uvalue:
            return
        dl = None
        ovalue = value  # Original value
        if "<" in uvalue:
            # import pudb; pu.db
            dl = value.lstrip("<").strip()
            value = dl
        name, fieldname = names
        name = name.strip()
        mo = COMPRE.match(name)
        add = self.g.add
        rel = PT[normURI(name)]

        def finish():
            if self.sample and not delim:
                # print(type(ovalue))
                # print("->{}->{}".format(rel, repr(ovalue)))
                add((self.sample, rel, Literal(ovalue)))

        def degs(v):
            if isinstance(v, float):
                return v
            # 107°25'28.48"В
            m = DEGRE.match(v)
            if m is None:
                return v
            d,_,m,_,s,_,_,dir = m.groups()
            d,m,s = [float(v) for v in [d,m,s]]
            d += m/60.
            d += s/3600.
            return d

        def unit(m, rupper):
            if 'PPM' in rupper:
                add((m, PT.unit, PPM))
            elif 'INT' in rupper:
                add((m, PT.unit, P.Int))
            elif '%' in fieldname:
                add((m, PT.unit, Percent))
            else:
                add((m, PT.unit, P.UnknowUnit))
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        if name in ['ППП', 'ппп']:
            rel = PT.il  # ignition losses
            m = BNode()
            add((self.sample, PT.measure, m))
            add((m, PT.value, Literal(value)))
            add((m, RDF.type, GeoMeasure))
            add((m, RDF.type, IgnitionLosses))
            u = name.upper()
            unit(m, u)
            return

        if mo is None:
            if name == 'с_ш':
                rel = Lat
                ovalue = degs(ovalue)
            if name == 'в_д':
                rel = Long
                ovalue = degs(ovalue)
            if name in ['sku', 'номер']:
                rel = SDO.sku
                if isinstance(ovalue, float):
                    ovalue=int(ovalue)
            finish()
            return
        else:
            comp = mo.group(1)
            rest = mo.group(3)
            # print(name, comp, rest, mo.groups(), fieldname)
            rc = COMPELRE.findall(comp)
            el1 = rc[0]
            el = ELRE.match(el1).group(1)
            eliri = elem(el)

        if eliri is None:  # This is not a compound
            finish()
            return

        def make_detlim():
            m = BNode()
            add((self.dsiri, PT.detectionLimit, m))
            add((m, RDF.type, PT.DetectionLimit))
            return m

        m = None
        # import pudb; pu.db
        # if name == 'Sn_PPM':
        #     print(name, value)
        #     import pudb; pu.db

        if self.sample and not delim:
            m = BNode()
            add((self.sample, PT.measure, m))
            add((m, RDF.type, GeoMeasure))
        if delim:
            m = make_detlim()

        rupper = rest.upper()

        if dl is None:
            add((m, PT.value, Literal(value)))
            unit(m, rupper)

        def finish_dl(e, m):
            if delim:
                self.dlims[e] = m
            if dl is not None:
                dlm = self.dlims.get(e, None)
                if dlm is not None:
                    add((m, PT.value, dlm))
                else:
                    # m is description, connected to a sample
                    md = make_detlim()
                    add((md, PT.value, Literal(value)))
                    add((md, RDF.type, GeoMeasure))
                    unit(md, rupper)
                    self.dlims[e] = md
                    add((m, PT.value, md))

        if len(rc) > 1 or el1 != el:  # Compound, e.g. oxide
            compname = 'compound-'+normURI(comp)
            cb = COMPOUNDS.get(compname, None)
            if cb is None:
                cb = PT[compname]
                add((cb, RDF.type, PT.Compound))
                add((cb, PT.Formula, Literal(comp)))
                add((cb, RDFS.label, Literal(comp)))
                COMPOUNDS[compname] = cb
            add((m, PT.compound, cb))

            finish_dl(comp, m)

        elif len(rc) == 1 and el1 == el and eliri is not None:
            add((m, MT.element, eliri))
            finish_dl(el, m)
        else:
            print("#!ERROR unknown combination of {}, and {}=?={}: {}.".format(rc,
                                                                                el1,
                                                                                el,
                                                                                eliri))
            quit()
        if "TOT" in rupper or "ОБЩ" in rupper:
            add((m, PT.total, Literal(True)))

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

        detlim = self.state == State.DETLIM
        if self.data() or detlim:
            if self.sample is None:
                self.c(row[self.sample_col], rx, self.sample_col, detlim)
            for i, cell in enumerate(row):
                self.c(cell, rx, i, detlim)
            return

        if self.state == State.HEADER:
            for i, cell in enumerate(row):
                self.h(cell, rx, i)
            return

        if self.state == State.CLASS:
            self.prev_class = None
            for i, cell in enumerate(row):
                self.hc(cell, rx, i)
            return

        if self.state == State.LOCATION:
            for cell in row:
                if cell.ctype == xlrd.XL_CELL_TEXT:
                    if len(self.locations) > self.loc_fence:
                        self.locations.pop()
                    loc = cell.value.strip()
                    self.proc_loc(loc)

    def c(self, cell, row, col, detlim=False):
        if cell.ctype in [xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK]:
            return
        try:
            field, fieldname = self.header[col]
        except KeyError as k:
            print("#! ERROR header key {} not in {} row: {}".format(k, self.header, row))
            # quit()
            return

        prt = self.cls.get(col, '')
        if '%' in prt and 'PPM' not in field:
            prt = '_%'
        elif 'PPM' in prt and '%' not in field:
            prt = '_PPM'
        else:
            prt = ''

        add = self.g.add
        ds = self.dsiri
        if field == self._sample_iri_:
            val = cell.value
            if isinstance(val, str):
                name = val.replace(' ', '')
            else:
                if isinstance(val, float):
                    if val.is_integer():
                        name = str(int(val))
                    else:
                        name = "{}".format(val)
                else:
                    name = "{}".format(val)
            self.sample = P[name]
            add((ds, self._sample_iri_, self.sample))
            add((self.sample, RDF.type, GeoSample))
            add((self.sample, RDF.type, SpatialThing))
            self.belongs(self.sample)
        elif self.sample is not None:
            self.proc_comp((field+prt, fieldname+prt), cell.value)
        elif detlim:
            self.proc_comp((field+prt, fieldname+prt), cell.value, detlim)
        else:
            print('#! ERROR: nowhere to store {} R:{} C:{}\n#!{}'.format(cell,
                                                                         row,
                                                                         col,
                                                                         self.header))
            quit()

    def h(self, cell, row, col):
        if cell.ctype in [xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK]:
            return
        orig = str(cell.value)
        name = normURI(orig)
        if name in self._sample_names_:
            name = self._sample_iri_
            self.sample_col = col
        self.header[col] = (name, orig)

    def hc(self, cell, row, col):
        if cell.ctype in [xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK]:
            if self.prev_class is None:
                return
            orig, name = self.prev_class
        else:
            orig = str(cell.value).replace('мг/кг', 'PPM')
            name = normURI(orig)
        self.cls[col] = orig
        self.prev_class = orig, name

class Yarki(ImpState):
    _sample_names_ = ['Sample', 'Полевой', 'sample']

class Kharantsy(ImpState):
    _sample_names_ = ['Номер_пробы']

class Khuzhir(Kharantsy):
    pass


FILES = {
         'Данные бар Ярки Сев Байкал.xls': (Yarki, ['Северный Байкал', 'Ярки']),
         'Данные Харанцы Ольхон.xls': (Kharantsy, ['Ольхон', 'Харанцы']),
         'Данные Хужир Ольхон.xls': (Khuzhir, ['Ольхон', 'Хужир']),
         'Сводная таблица РФА_Бураевская площадь № 70-2024.xls': (Yarki, ['Бураевская площадь']),
         'Сводная таблица РФА-2022 Усть-Кут.xls': (Yarki, ['Усть-Кутская площадь']),
        }


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


if __name__ == "__main__":
    for file, comp in FILES.items():
        parse_xl(file, comp)
    with open(TARGET, "w") as o:
        o.write(G.serialize(format='turtle'))
    with open(TARGETMT, "w") as o:
        o.write(GMT.serialize(format='turtle'))
    print("#!INFO: Normal exit")
