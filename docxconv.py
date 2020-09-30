from glob import glob
import re
import os
import win32com.client as win32
from win32com.client import constants

# Create list of paths to .doc files
# path = os.getcwd() + '\\BOMWordCopies\\003GMP.doc'


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


save_as_docx("003GENERIC_bom.doc", "BOMWordCopies")
