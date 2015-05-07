'''
Created on Jul 27, 2014

@author: vvaka
'''
from docx import Document

document = Document('dwdm_template.docx')
document.save('dwdm_templateUpdated.docx')
sections = document.sections
print len(sections)
for section in sections:
  print(section.start_type)
document.add_table(3, 4)