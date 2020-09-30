import os
import re
import win32com.client as win32
from win32com.client import constants


class ExtractBOM:
    @staticmethod
    def renameBOM(doc):
        stopwords = ['R00', 'R01', 'R01q', 'R02', 'R03']
        new_name = re.split('_|\.|\s', doc)

        result_words = ''.join([word for word in new_name if word not in stopwords])
        return ''.join([i for i in result_words if not i.isalpha()])

    # TODO: Improve this?
    def moveDocs(self, doc, oldloc, loc, ext):
        bom_number = self.renameBOM(doc)
        try:
            os.rename(oldloc + doc, loc + f"{bom_number}" + ext)
        except WindowsError:
            os.remove(loc + f"{bom_number}" + ext)
            os.rename(oldloc + doc, loc + f"{bom_number}" + ext)
        self.save_as_docx(bom_number + ext, loc)
        os.remove(loc + f"{bom_number}" + ext)

    def convertDocx(self, doc, loc):
        self.moveDocs(doc, loc + "/", loc + "/", "GMP.doc")

    @staticmethod
    def save_as_docx(doc, loc):
        # Opening MS Word
        path = os.getcwd() + '\\' + loc + '\\' + doc
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(path)
        doc.Activate()

        # Rename path with .docx
        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

        # Save and Close
        word.ActiveDocument.SaveAs(
            new_file_abs, FileFormat=constants.wdFormatXMLDocument
        )
        doc.Close(False)
