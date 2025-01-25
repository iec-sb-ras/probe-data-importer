from SPARQLWrapper import SPARQLWrapper, POST

sparql = SPARQLWrapper("http://ktulhu.isclan.ru:8890/sparql")
# sparql.setCredentials("some-login", "some-password") # if required
sparql.setMethod(POST) # this is the crucial option

sparql.setQuery("""
    prefix rdfs: <http://www.w3.org/2000/01/rdf-schema#>
    select distinct ?Name
    from <http://www.daml.org/2003/01/periodictable/PeriodicTable#>
    where {[] rdfs:label ?Name} LIMIT 100
    """
)

results = sparql.query()
print(results.response.read().decode('utf-8'))
