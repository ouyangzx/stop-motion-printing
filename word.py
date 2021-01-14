from win32com.client import Dispatch

# word = Dispatch('Word.Application')
# word = Dispatch('wps Application')
# word = Dispatch('kwps Application')
word = Dispatch('word.Application')
word.Visible = True


test_docx = 'C:\\Users\\Administrator\\IdeaProjects\\stop-motion-printing\\test1.docx'
doc = word.Documents.Open(test_docx)
