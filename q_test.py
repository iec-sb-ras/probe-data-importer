from common import NTQuery, SAMPLEGRAPH, quicktest
from namespace import PT, P, MT
from collections import namedtuple

def pollution_data(site, element):
    samples = """
    # SELECT *
    SELECT ?uri, ?long, ?lat
    FROM <{graph}>
    WHERE {{
       ?site rdfs:label ?site_name .
       ?site a pt:Site  .
       FILTER (?site_name = "{site}"@ru)
       ?uri pt:location ?site .
       ?uri a pt:Sample .
       ?uri rdfs:label ?sample_name .
       ?uri wgs:location ?loc .
       ?loc wgs:lat ?lat .
       ?loc wgs:long ?long .
    }}
    """

    qs = NTQuery(samples, SAMPLEGRAPH, site=site)

    element_data = """
    SELECT ?value, ?unit
    FROM <{graph}>
    WHERE {{
        <{sample}> a pt:Sample .
        <{sample}> pt:measurement ?m .
        ?m mt:element {element} .
        ?m pt:value ?value .
        ?m pt:unit ?u .
        OPTIONAL {{
           ?u rdfs:label ?unit .
        }}
    }}
    """

    qs.print()
    for sample in qs.results():
        qe = NTQuery(element_data, SAMPLEGRAPH, sample=sample.uri,
                     element=element.n3())
        for el in qe.results():
            row = namedtuple('Row', 'sample measurement')
            yield row(sample, measurement=el)
    #     else:
    #         print("No measurements for {} in sample {}".format(site, sample))
    # else:
    #     print("No samples for {}".format(site))


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

if __name__ == "__main__":
    for row in pollution_data("Харанцы", MT.Ga):
        print(row)
