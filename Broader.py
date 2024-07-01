import pandas as pd
from rdflib import Graph, URIRef, Literal, RDF, Namespace
from rdflib.namespace import RDFS, OWL, SKOS, DCTERMS, XSD

# Define the expected column names
expected_columns = [
    'Namespace', 'rdf:Description rdf:about', 'rdf:type rdf:resource', 'rdf:type rdf:resource.1',
    'rdfs:subClassOf rdf:resource', 'rdfs:subClassOf rdf:resource.1', 'rdfs:subClassOf rdf:resource.2',
    'skos:altLabel xml:lang="en"', 'skos:prefLabel xml:lang="en"', 'skos:definition xml:lang="en"',
    'skos:note xml:lang="en"', 'dct:source xml:lang="en"', 'skos:exactMatch',
    'dct:created rdf:datatype="http://www.w3.org/2001/XMLSchema#date"', 'dct:contributor',
    'rdfs:isDefinedBy rdf:resource', 'skos:note xml:lang="en".1', 'skos:note xml:lang="en".2', 'skos:broader'
]

# Load Excel data
excel_file = r'C:\Users\jenni\NI2Automation\Broader Match Terms.xlsx'
print(f"Loading Excel file from {excel_file}")
df = pd.read_excel(excel_file, header=1, names=expected_columns)  # Set header row and column names
print("Excel file loaded successfully")

# Print the column names to debug
print("Columns in the DataFrame:", df.columns)

# Create RDF graph
g = Graph()

# Define namespaces
nis2v = Namespace("http://JP_ontology.org/nis2v/")
iso = Namespace("http://JP_ontology.org/iso27001/")
dct = Namespace("http://purl.org/dc/terms/")
g.bind("nis2v", nis2v)
g.bind("iso27001", iso)
g.bind("rdfs", RDFS)
g.bind("owl", OWL)
g.bind("skos", SKOS)
g.bind("dct", dct)
g.bind("dcterms", DCTERMS)

print("Namespaces defined and bound")

# Function to construct the full IRI for terms
def construct_iri(value):
    if pd.notna(value) and not value.startswith("http://"):
        return URIRef(nis2v[value])
    return URIRef(value) if pd.notna(value) else None

# Dictionary to track processed URIs
processed_uris = {}

# Iterate through the rows of the DataFrame and add triples to the graph
for index, row in df.iterrows():
    print(f"Processing row {index + 1}")
    print("Row data:", row)
    class_uri = URIRef(nis2v[row['rdf:Description rdf:about']])

    if class_uri in processed_uris:
        # If the term has already been processed, add the new relations to the existing term
        if pd.notna(row['rdfs:subClassOf rdf:resource']):
            g.add((class_uri, RDFS.subClassOf, construct_iri(row['rdfs:subClassOf rdf:resource'])))
        if pd.notna(row['rdfs:subClassOf rdf:resource.1']):
            g.add((class_uri, RDFS.subClassOf, construct_iri(row['rdfs:subClassOf rdf:resource.1'])))
        if pd.notna(row['rdfs:subClassOf rdf:resource.2']):
            g.add((class_uri, RDFS.subClassOf, construct_iri(row['rdfs:subClassOf rdf:resource.2'])))
        if pd.notna(row['skos:note xml:lang="en"']):
            g.add((class_uri, SKOS.note, Literal(row['skos:note xml:lang="en"'], lang='en')))
        if pd.notna(row['skos:note xml:lang="en".1']):
            g.add((class_uri, SKOS.note, Literal(row['skos:note xml:lang="en".1'], lang='en')))
        if pd.notna(row['skos:note xml:lang="en".2']):
            g.add((class_uri, SKOS.note, Literal(row['skos:note xml:lang="en".2'], lang='en')))
        if pd.notna(row['skos:broader']):
            g.add((class_uri, SKOS.broader, URIRef(row['skos:broader'])))
    else:
        # Process the term normally if it hasn't been encountered before
        g.add((class_uri, RDF.type, URIRef(row['rdf:type rdf:resource'])))
        g.add((class_uri, RDF.type, URIRef(row['rdf:type rdf:resource.1'])))

        if pd.notna(row['rdfs:subClassOf rdf:resource']):
            g.add((class_uri, RDFS.subClassOf, construct_iri(row['rdfs:subClassOf rdf:resource'])))
        if pd.notna(row['rdfs:subClassOf rdf:resource.1']):
            g.add((class_uri, RDFS.subClassOf, construct_iri(row['rdfs:subClassOf rdf:resource.1'])))
        if pd.notna(row['rdfs:subClassOf rdf:resource.2']):
            g.add((class_uri, RDFS.subClassOf, construct_iri(row['rdfs:subClassOf rdf:resource.2'])))

        if pd.notna(row['skos:altLabel xml:lang="en"']):
            g.add((class_uri, SKOS.altLabel, Literal(row['skos:altLabel xml:lang="en"'], lang='en')))
        if pd.notna(row['skos:prefLabel xml:lang="en"']):
            g.add((class_uri, SKOS.prefLabel, Literal(row['skos:prefLabel xml:lang="en"'], lang='en')))
        if pd.notna(row['skos:definition xml:lang="en"']):
            g.add((class_uri, SKOS.definition, Literal(row['skos:definition xml:lang="en"'], lang='en')))
        if pd.notna(row['skos:note xml:lang="en"']):
            g.add((class_uri, SKOS.note, Literal(row['skos:note xml:lang="en"'], lang='en')))
        if pd.notna(row['dct:source xml:lang="en"']):
            g.add((class_uri, DCTERMS.source, Literal(row['dct:source xml:lang="en"'], lang='en')))
        if pd.notna(row['skos:exactMatch']):
            g.add((class_uri, SKOS.exactMatch, URIRef(row['skos:exactMatch'])))
        if pd.notna(row['dct:created rdf:datatype="http://www.w3.org/2001/XMLSchema#date"']):
            g.add((class_uri, DCTERMS.created, Literal(row['dct:created rdf:datatype="http://www.w3.org/2001/XMLSchema#date"'], datatype=XSD.date)))
        if pd.notna(row['dct:contributor']):
            g.add((class_uri, DCTERMS.contributor, Literal(row['dct:contributor'])))
        if pd.notna(row['rdfs:isDefinedBy rdf:resource']):
            g.add((class_uri, RDFS.isDefinedBy, URIRef(row['rdfs:isDefinedBy rdf:resource'])))
        if pd.notna(row['skos:note xml:lang="en".1']):
            g.add((class_uri, SKOS.note, Literal(row['skos:note xml:lang="en".1'], lang='en')))
        if pd.notna(row['skos:note xml:lang="en".2']):
            g.add((class_uri, SKOS.note, Literal(row['skos:note xml:lang="en".2'], lang='en')))
        if pd.notna(row['skos:broader']):
            g.add((class_uri, SKOS.broader, URIRef(row['skos:broader'])))

        # Mark this URI as processed
        processed_uris[class_uri] = True

    print(f"Added triples for row {index + 1}")

# Serialize graph to RDF/XML
rdf_file = r'C:\Users\jenni\OneDrive\Documents\UCD Cybersecurity Masters\COMP47830 - Cybersecurity Research Project\Protege\Automation-Output\Broader_Match_Terms_v1.rdf'
print(f"Serializing graph to {rdf_file}")
g.serialize(destination=rdf_file, format='xml')

print(f"RDF data has been saved to {rdf_file}")
