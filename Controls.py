import pandas as pd
from rdflib import Graph, URIRef, Literal, RDF, Namespace
from rdflib.namespace import RDFS, OWL, SKOS, DCTERMS, XSD

# Check if openpyxl is available
try:
    import openpyxl
    print("openpyxl is available")
except ImportError:
    print("openpyxl is not available")

# Load Excel data
excel_file = r'C:\Users\jenni\NI2Automation\Class Automation - Controls.xlsx'  # Ensure the correct path
print(f"Loading Excel file from {excel_file}")
df = pd.read_excel(excel_file)
print("Excel file loaded successfully")

# Create RDF graph
g = Graph()

# Define namespaces
nis2v = Namespace("http://JP_ontology.org/nis2v/")
g.bind("nis2v", nis2v)
g.bind("rdfs", RDFS)
g.bind("owl", OWL)
g.bind("skos", SKOS)
g.bind("dcterms", DCTERMS)

print("Namespaces defined and bound")

# Function to construct the full IRI for rdfs:subClassOf
def construct_iri(value):
    if pd.notna(value) and not value.startswith("http://"):
        return URIRef(nis2v[value])
    return URIRef(value) if pd.notna(value) else None

# Iterate through the rows of the DataFrame and add triples to the graph
for index, row in df.iterrows():
    print(f"Processing row {index+1}")
    class_uri = URIRef(nis2v[row['rdf:Description rdf:about']])
    g.add((class_uri, RDF.type, URIRef(row['rdf:type rdf:resource'])))
    g.add((class_uri, RDF.type, URIRef(row['rdf:type rdf:resource.1'])))

    if pd.notna(row['rdfs:subClassOf rdf:resource']):
        g.add((class_uri, RDFS.subClassOf, construct_iri(row['rdfs:subClassOf rdf:resource'])))

    if pd.notna(row['rdfs:subClassOf rdf:resource.1']):
        g.add((class_uri, RDFS.subClassOf, construct_iri(row['rdfs:subClassOf rdf:resource.1'])))

    if pd.notna(row['skos:altLabel xml:lang="en"']):
        alt_label_text = row['skos:altLabel xml:lang="en"'].replace('http://JP_ontology.org/nis2v/', '')
        g.add((class_uri, SKOS.altLabel, Literal(alt_label_text, lang='en')))

    if pd.notna(row['skos:prefLabel xml:lang="en"']):
        g.add((class_uri, SKOS.prefLabel, Literal(row['skos:prefLabel xml:lang="en"'], lang='en')))

    if pd.notna(row['skos:definition xml:lang="en"']):
        g.add((class_uri, SKOS.definition, Literal(row['skos:definition xml:lang="en"'], lang='en')))

    if pd.notna(row['dct:source xml:lang="en"']):
        g.add((class_uri, DCTERMS.source, Literal(row['dct:source xml:lang="en"'], lang='en')))

    if pd.notna(row['dct:created rdf:datatype="http://www.w3.org/2001/XMLSchema#date"']):
        g.add((class_uri, DCTERMS.created, Literal(row['dct:created rdf:datatype="http://www.w3.org/2001/XMLSchema#date"'], datatype=XSD.date)))

    if pd.notna(row['dct:contributor']):
        g.add((class_uri, DCTERMS.contributor, Literal(row['dct:contributor'])))

    if pd.notna(row['rdfs:isDefinedBy rdf:resource']):
        g.add((class_uri, RDFS.isDefinedBy, URIRef(row['rdfs:isDefinedBy rdf:resource'])))

    print(f"Added triples for row {index+1}")

# Serialize graph to RDF/XML
rdf_file = r'C:\Users\jenni\OneDrive\Documents\UCD Cybersecurity Masters\COMP47830 - Cybersecurity Research Project\Protege\Automation-Output\Auto_Terms1.rdf'
print(f"Serializing graph to {rdf_file}")
g.serialize(destination=rdf_file, format='xml')

print(f"RDF data has been saved to {rdf_file}")
