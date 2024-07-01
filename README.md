# NIS2OntologyAutomation
This repository contains scripts to automate the creation of an ontology for the NIS2 Regulatory Assessment Tool. These scripts process Excel data, generate RDF graphs, and serialize the data into RDF/XML format.

# Features
Automated Ontology Creation: Dynamically converts Excel data into RDF graphs for semantic analysis.
Namespace Management: Defines and manages multiple namespaces for the RDF data.

# Files and Directories
Lib: Directory containing library files used by the scripts.
Scripts: Directory containing additional scripts.
.gitignore: Specifies files and directories to be ignored by git.
Answers.py: Script for handling answers related to ontology terms.
Broader.py: Script for processing broader match terms.
Complex.py: Script for handling complex terms in the ontology.
Controls.py: Script for managing control terms in the ontology.
ExactMatch.py: Script for processing exact match terms.
Narrower.py: Script for handling narrower match terms.
NewTerms.py: Script for adding new terms to the ontology.
NIS2V Terms.py: Script for processing various terms related to the NIS2 ontology.
pyvenv.cfg: Python virtual environment configuration file.
Question.py: Script for handling questions related to ontology terms.

# Getting Started
> Prerequisites
Python 3.x
Required Python packages (install via pip):
pandas
rdflib

> Installation
Clone the repository:
git clone https://github.com/yourusername/NIS2OntologyAutomation.git
cd NIS2OntologyAutomation

Install the required packages:
pip install pandas rdflib

# Usage
Load and Process Excel Data: Place your Excel files in the specified paths as referenced in the scripts.
Ensure your Excel files are correctly formatted and located at the specified paths in the scripts. 
The generated RDF files will be saved in the specified output directories.
