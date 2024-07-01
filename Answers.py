import pandas as pd
from rdflib import Graph, URIRef, Literal, RDF, Namespace
from rdflib.namespace import RDFS, OWL, SKOS, DCTERMS, XSD

def replace_first_dot_with_underscore(text):
    # Find the position of the first full stop in the text
    dot_position = text.find('.')

    # If there is a full stop, replace the first one with an underscore
    if dot_position != -1:
        text = text[:dot_position] + '' + text[dot_position + 1:]

    return text

# Load Excel data
excel_file = r'C:\Users\jenni\NI2Automation\Class Automation - Answers.xlsx'  # Ensure the correct path
print(f"Loading Excel file from {excel_file}")
df = pd.read_excel(excel_file)
print("Excel file loaded successfully")

# Print the column names to debug
print("Columns in the DataFrame:", df.columns)

# Fill in blank cells in 'rdf:type rdf:resource' based on 'MCQ.xAnswerA'
for i in range(len(df)):
    if pd.isna(df.at[i, 'rdf:type rdf:resource']):
        # Look for the nearest non-blank value above the current row
        j = i - 1
        while pd.isna(df.at[j, 'rdf:type rdf:resource']) and j >= 0:
            j -= 1
        if j >= 0:
            df.at[i, 'rdf:type rdf:resource'] = df.at[j, 'rdf:type rdf:resource']

# Update the rdf:Description rdf:about column
df['rdf:Description rdf:about'] = df['rdf:Description rdf:about'].apply(replace_first_dot_with_underscore)
df['rdf:Description rdf:about'] = 'http://JP_ontology.org/nis2v/' + df['rdf:Description rdf:about']

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
    class_uri = URIRef(row['rdf:Description rdf:about'])
    g.add((class_uri, RDF.type, URIRef(nis2v[row['rdf:type rdf:resource']])))

    if pd.notna(row['skos:prefLabel']):
        g.add((class_uri, SKOS.prefLabel, Literal(row['skos:prefLabel'], lang='en')))

    if pd.notna(row['skos:altLabel']):
        g.add((class_uri, SKOS.altLabel, Literal(row['skos:altLabel'], lang='en')))

    if pd.notna(row['skos:definition xml:lang="en"']):
        g.add((class_uri, SKOS.definition, Literal(row['skos:definition xml:lang="en"'], lang='en')))

    print(f"Added triples for row {index + 1}")

# Serialize graph to RDF/XML
rdf_file = r'C:\Users\jenni\OneDrive\Documents\UCD Cybersecurity Masters\COMP47830 - Cybersecurity Research Project\Protege\Automation-Output\Auto_Answers.rdf'
print(f"Serializing graph to {rdf_file}")
g.serialize(destination=rdf_file, format='xml')

print(f"RDF data has been saved to {rdf_file}")
