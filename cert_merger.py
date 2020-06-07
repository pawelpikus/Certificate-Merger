from mailmerge import MailMerge
import pandas as pd
import sys

try:
    if sys.argv[1].endswith(('.doc', '.docx')):
        template_filename = sys.argv[1]
    else:
        print('Template must be in a .doc or .docx format.')
        sys.exit(0)
    if sys.argv[2].endswith(('.xls', '.xlsx')):
        data_filename = sys.argv[2]
    else:
        print('Data must be in a .xls or .xlsx format.')
        sys.exit(0)

    data = pd.read_excel(data_filename)
    #  change all df values to str, needed by merge_templates()
    data = data.applymap(str)
    data_dict = data.to_dict(orient='index')

    students = []
    for key in data_dict.keys():
        students.append(data_dict[key])

    with MailMerge(template_filename) as document:
        document.merge_templates(students, separator='page_break')
        document.write('certificates.docx')
    print('Done.')
except IndexError:
    print("Type: cert_merger.py [arg1=template] [arg2=data]")
except FileNotFoundError as err:
    print(err)

