#!/bin/env python
from common import NTQuery, SAMPLEGRAPH, quicktest
from rdflib import URIRef
from namespace import PT, P, MT
from collections import namedtuple
import pprint
import csv
import sys

query_sample_data = """
    # SELECT *
    SELECT ?uri, ?site_name, ?sample_name,
           # ?long, ?lat,
           ?element, ?value, ?unitid, ?unit
    FROM <{graph}>
    WHERE {{
       ?main_site rdfs:label ?m_site_name .
       ?main_site a pt:Site  .
       FILTER (?m_site_name = "{site}"@ru)
       ?main_site ^pt:location* ?site .
       ?site rdfs:label ?site_name .
       ?site a pt:Site  .
       ?uri pt:location ?site .
       ?uri a pt:Sample .
       ?uri rdfs:label ?sample_name .
       # ?uri wgs:location ?loc .
       # ?loc a wgs:Point .
       # ?loc wgs:lat ?lat .
       # ?loc wgs:long ?long .
       ?uri pt:measurement ?m .
       # ?m mt:element {element} .
       ?m mt:element ?element .
       ?m pt:value ?value .
       ?m pt:unit ?unitid .
       OPTIONAL {{
          ?unitid rdfs:label ?unit .
       }}
    }}
    """

query_sample_data_with_coords = """
    select ?slab ?el ?val ?unit ?lat ?long
    FROM <{graph}>
    where {
      ?s a pt:Sample .
      { ?s pt:location pi:Харанцы . } union { ?s pt:location pi:Хужир . }
      ?s wgs:location ?loc .
      ?s rdfs:label ?slab .
      ?loc  a wgs:Point .
      ?loc wgs:lat ?lat .
      ?loc wgs:long ?long .
      ?s pt:measurement ?m .
      ?m a pt:Measurement .
      ?m pt:unit ?unit .
      # ?unit rdfs:label ?unitlab .
      ?m pt:value ?val .
      ?m mt:element ?el .
    }
"""


def pollution_data(query, site, element=None):
    samples = query

    if element is not None:
        qs = NTQuery(samples, SAMPLEGRAPH, site=site, element=element.n3())
    else:
        qs = NTQuery(samples, SAMPLEGRAPH, site=site, element=None)

    # qs.print()
    return qs.results()


# quicktest("""
#     SELECT *
#     FROM <{graph}>
#     WHERE {{
#         <http://crust.irk.ru/ontology/pollution/1.0/AGS-0110> a pt:Sample .
#         <http://crust.irk.ru/ontology/pollution/1.0/AGS-0110> pt:measurement ?m .
#         ?m mt:element <http://www.daml.org/2003/01/periodictable/PeriodicTable#Ga> .
#         # ?m mt:element ?el .
#         ?m pt:value ?value .
#     }}
# """, SAMPLEGRAPH)

# quicktest(query_sample_data, SAMPLEGRAPH, site="Бураевская площадь",
#          element = MT.Ga.n3(), debug=True)

if __name__ == "__main__":
    tbl = {}
    headers = set(['sample', 'site'])

    def u(s):
        if s.startswith("http"):
            return URIRef(s)
        return s
#    for row in pollution_data(query_sample_data,
#                              "Бураевская площадь", None):

    for row in pollution_data(query_sample_data, sys.argv[1], None):
        sample = u(row.uri)
        element = u(row.element)
        headers.add(element)
        s = tbl.setdefault(sample, {
            "sample": row.sample_name,
            "site": row.site_name
        })
        try:
            s[element] = float(row.value)
        except ValueError:
            s[element] = 0.0
        # Row(uri='http://crust.irk.ru/ontology/pollution/1.0/2464-1', site_name='Ивановский', sample_name='2464-1', element='http://www.daml.org/2003/01/periodictable/PeriodicTable#V', value='4.2999999999999998224', unitid='http://crust.irk.ru/ontology/pollution/terms/1.0/PPM', unit='мг/кг')

    with open(sys.argv[2], "w") as o:
        wr = csv.writer(o, quotechar='"', quoting=csv.QUOTE_STRINGS)

        def fr(row):
            for e in row:
                # print(e, type(e))
                if isinstance(e, URIRef):
                    yield e.fragment
                else:
                    yield e

        h = fr(headers)
        # wr.writerow(headers)
        wr.writerow(h)

        def g(headers, row):
            for h in headers:
                if h in row:
                    yield row[h]
                else:
                    yield 0.0

        for sample, row in tbl.items():
            wr.writerow(g(headers, row))
