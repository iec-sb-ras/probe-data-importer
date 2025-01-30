from rdflib import Namespace, BNode, URIRef, RDFS, Graph
from rdflib.namespace import WGS
import csv
from collections import namedtuple


PT = Namespace('http://crust.irk.ru/ontology/pollution/terms/1.0/')
P = Namespace('http://crust.irk.ru/ontology/pollution/1.0/')
MT = Namespace('http://www.daml.org/2003/01/periodictable/Periodic Table#')

G = Graph(bind_namespaces="rdflib")


for _ in [G]:
    _.bind('pt', PT)
    _.bind('pi', P)
    _.bind('mt', MT)


def imp(stream):
    c = csv.reader(stream)
    headers = next(c)
    nt = namedtuple('Tuple', headers)
    for row in c:
        # d = dict(zip(headers, row))
        # print(d)
        o = nt(*row)




FILES = "t.csv  u-k.csv".split()
FIELDS = ['', 'Сводная_таблица_РФА_2022_Усть_Кут, XSЛист1', ]


if __name__=="__main__":
    for f in FILES:
        with open(f) as i:
            imp(i)
