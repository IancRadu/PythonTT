import docxtpl
from docxtpl import DocxTemplate

tpl = DocxTemplate('your_template.docx')
rt = docxtpl.RichText('You can add an hyperlink, here to ')

info_to_replace = {'rt': "'google', url_id=tpl.build_url_id('http://google.com')"}
tpl.render(info_to_replace)
# Save and create the file in the location and with the name specified between ()
tpl.save('result.docx')
