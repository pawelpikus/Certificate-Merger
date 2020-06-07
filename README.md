## Certificate Merger

This Python script automates the process of filling-in names, scores, and other fields in certificates/diplomas or any other Word documents.
It's features include:

 - supporting `.doc`, `.docx` for certificate templates
 - supporting `.xls`, `.xlsx` for data files 
 - merging any number of documents 
 - simple command-line execution
 - output to a single `certificates.docx` file



# Installation
The script is written in Python 3.8. Tested od Windows 10. 

Dependencies:
- pandas 1.0.4
`pip install pandas`
- docx-mailmerge 0.5.0
`pip install docx-mailmerge`

## Usage

**Template** must be a `.doc` or `.docx` file. For detailed instructions on how to prepare a template in a MS Word document using MergeField, see:

 - `docx-mailmerge` module documentation
 - A nice article with example code at: http://pbpython.com/python-word-template.html

**Data** must be a `.xls` or `.xlsx` file. The columns should have the same names as the MergeFields in the template.  

### CLI
In Terminal/Command Line, run: 
`py cert_maker.py path_to_file\template_file.docx path_to_file\data_file.xlsx`
