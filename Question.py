import pandas as pd
from rdflib import Graph, URIRef, Literal, RDF, Namespace
from rdflib.namespace import RDFS, OWL, SKOS, DCTERMS, XSD

# Load Excel data
excel_file = r'C:\Users\jenni\NI2Automation\Class Automation - Questions.xlsx'  # Ensure the correct path
print(f"Loading Excel file from {excel_file}")
df = pd.read_excel(excel_file)
print("Excel file loaded successfully")

# Print the column names to debug
print("Columns in the DataFrame:", df.columns)

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

# Iterate through the rows of the DataFrame and add triples to the graph
for index, row in df.iterrows():
    print(f"Processing row {index + 1}")
    print("Row data:", row)
    class_uri = URIRef(nis2v[row['skos:altLabel xml:lang="en"']])
    g.add((class_uri, RDF.type, URIRef(row['rdf:type rdf:resource'])))
    g.add((class_uri, RDF.type, URIRef(row['rdf:type rdf:resource.1'])))

    if pd.notna(row['rdfs:subClassOf rdf:resource']):
        g.add((class_uri, RDFS.subClassOf, URIRef(row['rdfs:subClassOf rdf:resource'])))

    if pd.notna(row['skos:prefLabel xml:lang"en"']):
        g.add((class_uri, SKOS.prefLabel, Literal(row['skos:prefLabel xml:lang"en"'], lang='en')))

    if pd.notna(row['skos:altLabel xml:lang="en"']):
        g.add((class_uri, SKOS.altLabel, Literal(row['skos:altLabel xml:lang="en"'], lang='en')))

    if pd.notna(row['skos:definition xml:lang="en"']):
        g.add((class_uri, SKOS.definition, Literal(row['skos:definition xml:lang="en"'], lang='en')))

    if pd.notna(row['skos:broader rdf:resource']):
        g.add((class_uri, SKOS.broader, URIRef(row['skos:broader rdf:resource'])))

    if pd.notna(row['dct:created rdf:datatype="http://www.w3.org/2001/XMLSchema#date"']):
        g.add((class_uri, DCTERMS.created, Literal(row['dct:created rdf:datatype="http://www.w3.org/2001/XMLSchema#date"'], datatype=XSD.date)))

    if pd.notna(row['dct:contributor']):
        g.add((class_uri, DCTERMS.contributor, Literal(row['dct:contributor'])))

    if pd.notna(row['rdfs:isDefinedBy rdf:resource']):
        g.add((class_uri, RDFS.isDefinedBy, URIRef(row['rdfs:isDefinedBy rdf:resource'])))

    if pd.notna(row['nis2v:hasAnswer rdf:resource']):
        g.add((class_uri, nis2v.hasAnswer, URIRef(row['nis2v:hasAnswer rdf:resource'])))

    if pd.notna(row['nis2v:hasAnswer rdf:resource.1']):
        g.add((class_uri, nis2v.hasAnswer, URIRef(row['nis2v:hasAnswer rdf:resource.1'])))

    if pd.notna(row['nis2v:hasAnswer rdf:resource.2']):
        g.add((class_uri, nis2v.hasAnswer, URIRef(row['nis2v:hasAnswer rdf:resource.2'])))

    if pd.notna(row['nis2v:relatedToControl rdf:resource']):
        g.add((class_uri, nis2v.relatedToControl, URIRef(row['nis2v:relatedToControl rdf:resource'])))

    if pd.notna(row['dct:subject rdf:resource']):
        g.add((class_uri, DCTERMS.subject, URIRef(row['dct:subject rdf:resource'])))

    print(f"Added triples for row {index + 1}")

# Serialize graph to RDF/XML
rdf_file = r'C:\Users\jenni\OneDrive\Documents\UCD Cybersecurity Masters\COMP47830 - Cybersecurity Research Project\Protege\Automation-Output\Auto_Questions.rdf'
print(f"Serializing graph to {rdf_file}")
g.serialize(destination=rdf_file, format='xml')

print(f"RDF data has been saved to {rdf_file}")
