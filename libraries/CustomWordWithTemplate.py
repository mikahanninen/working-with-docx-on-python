import olefile

from zipfile import ZipFile
from docxtpl import DocxTemplate
from robot.libraries.BuiltIn import BuiltIn


class CustomWordWithTemplate:
    def create_doc_with_template(self, template_file, outputfile, pdf_to_insert=None):
        BuiltIn().log_to_console(
            "SRC: %s\nTARGET:%s" % (pdf_to_insert["src"], pdf_to_insert["target"])
        )
        self.open_as_zip_file(template_file)
        doc = DocxTemplate(template_file)
        context = {"name": "Mika Hänninen"}
        doc.replace_embedded(pdf_to_insert["src"], pdf_to_insert["target"])
        doc.render(context)
        doc.save(outputfile)

    def open_as_zip_file(self, filename):
        pdf_count = 0
        try:
            BuiltIn().log_to_console(
                "Try to open the document as ZIP file: %s" % filename
            )
            with ZipFile(filename, "r") as zip:
                # Find files in the word/embeddings folder of the ZIP file
                for entry in zip.infolist():
                    BuiltIn().log_to_console(entry.filename)
                    if not entry.filename.startswith("word/embeddings/"):
                        continue

                    BuiltIn().log_to_console("Try to open the embedded OLE file")
                    with zip.open(entry.filename) as f:
                        if not olefile.isOleFile(f):
                            BuiltIn().log_to_console("Is NOT isoOleFile")
                            continue

                        ole = olefile.OleFileIO(f)

                        # CLSID for Adobe Acrobat Document
                        BuiltIn().log_to_console("clsid: %s" % ole.root.clsid)
                        if ole.root.clsid != "B801CA65-A1FC-11D0-85AD-444553540000":
                            BuiltIn().log_to_console("Is NOT Adode Acrobat Document")
                            continue

                        if not ole.exists("CONTENTS"):
                            BuiltIn().log_to_console("did NOT have CONTENTS")
                            continue

                        # Extract the PDF from the OLE file
                        pdf_data = ole.openstream("CONTENTS").read()

                        # Does the embedded file have a %PDF- header?
                        if pdf_data[0:5] == b"%PDF-":
                            pdf_count += 1

                            pdf_filename = "Document %d.pdf" % pdf_count

                            # Save the PDF
                            with open(pdf_filename, "wb") as output_file:
                                output_file.write(pdf_data)
        except:
            print("Unable to open '%s'" % filename)

    def transform_pdf_into_ole_object(self, filename):

        with open(filename, "rb") as f:
            ole = olefile.OleFileIO(f)
            BuiltIn().log_to_console("NO mut OLE !!")
            BuiltIn().log_to_console(dir(ole))