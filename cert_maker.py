from mailmerge import MailMerge
import pandas as pd
import sys

template_filename = sys.argv[1]
data_filename = sys.argv[2]

data = pd.read_excel(data_filename)
#  change all df values to str, needed by merge_templates()
data = data.applymap(str)
data_dict = data.to_dict(orient='index')

students = []
for key, value in data_dict.items():
    students.append(data_dict[key])

with MailMerge(template_filename) as document:
    document.merge_templates(students, separator='page_break')
    document.write('certificates.docx')
print('Done.')
