from win32com.client import Dispatch


class WpsWord:

    def __init__(self):
        self.word = Dispatch('word.Application')
        self.word.Visible = True

    def open(self,docx):
        self.word.Documents.Open(docx)
