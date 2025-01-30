from SPARQLWrapper import SPARQLWrapper, POST, JSON
import requests as rq
import os

try:
    del os.environ["HTTP_PROXY"]
    del os.environ["HTTPS_PROXY"]
except KeyError:
    pass
try:
    del os.environ["http_proxy"]
    del os.environ["https_proxy"]
except KeyError:
    pass

ENDPOINT = "http://ktulhu.isclan.ru:8890/sparql"
SAMPLEGRAPH = "http://localhost:8890/DAV/home/loader/rdf_sink/samples.ttl"

PREFIXES = """
    PREFIX pt: <http://crust.irk.ru/ontology/pollution/terms/1.0/>
    PREFIX p: <http://crust.irk.ru/ontology/pollution/1.0/>
    PREFIX mt: <http://www.daml.org/2003/01/periodictable/PeriodicTable#>
    PREFIX wgs: <https://www.w3.org/2003/01/geo/wgs84_pos#>
    PREFIX rdfs: <http://www.w3.org/2000/01/rdf-schema#>

"""

QUERY = PREFIXES + """
    select distinct ?Name
    from <http://www.daml.org/2003/01/periodictable/PeriodicTable#>
    where {[] rdfs:label ?Name} LIMIT 100
    """

sparql = SPARQLWrapper(ENDPOINT)
# sparql.setCredentials("some-login", "some-password") # if required
sparql.setMethod(POST) # this is the crucial option
sparql.setReturnFormat(JSON)

# sparql.setQuery(QUERY)

# results = sparql.query()
# print(results.response.read().decode('utf-8'))
# results.print_results()


def pollution_data(site):
    samples = PREFIXES + f"""
    SELECT *
    FROM <{SAMPLEGRAPH}>
    WHERE {{
       ?site rdfs:label ?site_name .
       ?site a pt:Site  .
       FILTER (?site_name = "{site}"@ru)
       ?sample pt:location ?site .
       ?sample a pt:Sample .
       ?sample rdfs:label ?sample_name .
    }}
    """


    print(samples)

    sparql.setQuery(samples)
    results = sparql.query()
    results.print_results()

if __name__=="__main__":
    pollution_data("Харанцы")
