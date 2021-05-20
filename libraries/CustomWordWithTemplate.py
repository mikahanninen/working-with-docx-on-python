from docxtpl import DocxTemplate

class CustomWordWithTemplate:

    def create_doc_with_template(self, template_file,  outputfile, pdf_to_insert=None):
        doc = DocxTemplate(template_file)
        context = { 'name' : "Mika Hänninen" }
        doc.render(context)
        doc.save(outputfile)