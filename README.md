## Certificate Merger

This Python script automates the process of filling-in names, scores, and other fields in certificates/diplomas or any other Word documents.
It's features include:

 - supporting `.doc`, `.docx` for certificate templates
 - supporting `.xls`, `.xlsx` for data files 
 - merging any number of documents using MergeField feature in Word
 - simple command-line execution
 - output to a single `certificates.docx` file

## Installation
The script is written in Python 3.8.
# Dependencies
- pandas 1.0.4
`pip install pandas`
- docx-mailmerge 0.5.0
`pip install docx-mailmerge`

##Usage

#

#CLI
`py cert_maker.py path_toFile\template_file.doc` path_toFIle\data.xlsx



