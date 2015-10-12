#coding:utf-8
import codecs 
import json

from docx import Document
from docx.shared import Inches


def replace():
	return ""

contents = {}
file_object = open("op_rec.json")
try:
	all_the_text = file_object.read()
	# 为了去除BOM 不得不做的检查。
	if all_the_text[:3] == codecs.BOM_UTF8:  
		all_the_text = all_the_text[3:] 
	#print all_the_text
	contents = json.loads( all_the_text )
except Exception as inst:
	print type(inst)     # the exception instance
	print inst.args      # arguments stored in .args
	print inst           # __str__ allows args to printed directly
      
finally:
	file_object.close( )

document = Document('op_rec_temp.docx')



# contents =  {'op_patient_name': u'张三',
# 			 'op_patient_sex': u'男',
# 			 'op_patient_age': u'40',
# 			 'op_department_name': u'肝胆外科'}

allkeys = contents.keys()

for paragraph in document.paragraphs:
	for aKey in allkeys:
	    replace_target = "{"+aKey+"}"
	    if replace_target in paragraph.text:
	    	aText = paragraph.text
	        print aText
	        bText = aText.replace( replace_target, contents.get(aKey))
	        print bText
	        paragraph.text = bText

document.add_picture('gdg-xian.png', width=Inches(2.5))

document.save('demo2.docx')
