from SPARQLWrapper import ( SPARQLWrapper, POST, JSON, CSV, RDF, RDFXML,
                            N3, JSONLD, XML )
import requests as rq
import os
import pprint

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
    PREFIX g: <{graph}>

"""

QUERY = (
    PREFIXES
    + """
    select distinct ?Name
    from <http://www.daml.org/2003/01/periodictable/PeriodicTable#>
    where {[] rdfs:label ?Name} LIMIT 100
    """
)

sparql = SPARQLWrapper(ENDPOINT)
# sparql.setCredentials("some-login", "some-password") # if required
sparql.setMethod(POST)  # this is the crucial option
sparql.setReturnFormat(JSON)
# sparql.setReturnFormat(XML)

# sparql.setQuery(QUERY)

# results = sparql.query()
# print(results.response.read().decode('utf-8'))
# results.print_results()


class Query:

    _prefixes_ = PREFIXES
    _endpoint_ = ENDPOINT

    def __init__(self, query, graphIRI, endpoint=None, **args):
        self.query = query
        self.graphIRI = graphIRI
        self.args = args
        if endpoint is None:
            endpoint = self._endpoint_
        self.endpoint = endpoint

    def results(self, debug=False):
        # import pudb; pu.db

        q = self._prefixes_+"\n\n"+self.query
        q = q.format(graph=self.graphIRI, **self.args)
        if debug:
            print(q)
        sparql = SPARQLWrapper(self.endpoint)
        sparql.setQuery(q)
        sparql.setMethod(POST)
        sparql.setReturnFormat(JSON)
        results = sparql.query()
        return conv(results.convert())


def pollution_data(site):
    samples = (
        """
    SELECT *
    FROM <{graph}>
    WHERE {{
       ?site rdfs:label ?site_name .
       ?site a pt:Site  .
       FILTER (?site_name = "{site}"@ru)
       ?sample pt:location ?site .
       ?sample a pt:Sample .
       ?sample rdfs:label ?sample_name .
    }}
    """
    )

    q = Query(samples, SAMPLEGRAPH, site=site)
    rc = q.results()
    pprint.pprint(rc)

def conv(results):
    header = results['head']['vars']
    binds = results['results']['bindings']
    rc = [{k: v['value'] for k, v in e.items()} for e in binds]
    return rc, header

if __name__ == "__main__":
    pollution_data("Харанцы")
