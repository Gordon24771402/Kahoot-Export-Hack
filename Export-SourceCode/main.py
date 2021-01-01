import string
import requests
from docx import Document
from docx.shared import Pt
import os
import time

URL = input('Copy and Paste Kahoot URL:\n\n')
answerURL = 'https://create.kahoot.it/rest/kahoots/{}/card/?includeKahoot=true'.format(URL.split('/')[-1])
data = requests.get(answerURL).json()['kahoot']['questions']
docQues = Document()
docAns = Document()

# Loop: Questions
for i in range(0, len(data)):
    text = '{ques}'.format(ques=data[i]['question'])
    # Remove &nbsp;
    text = text.replace('&nbsp;', '')
    # Paragraph of Questions
    paragraph = docQues.add_paragraph(text, style='List Number')
    # Format Options
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(8)

    # Loop: Choices
    for j in range(0, len(data[i]['choices'])):
        text = '\t{ansNum}. {ans}'.format(ansNum=list(string.ascii_uppercase)[j], ans=data[i]['choices'][j]['answer'])
        # Remove &nbsp;
        text = text.replace('&nbsp;', '')
        # Paragraph of Choices
        paragraph = docQues.add_paragraph(text)
        # Format Options
        paragraph.paragraph_format.line_spacing = 1.0
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(8)

    # Loop: Answers
    for j in range(0, len(data[i]['choices'])):
        ans = data[i]['choices'][j]['correct']
        # Determine Correctness
        if ans:
            # Paragraph of Choices
            text = list(string.ascii_uppercase)[j]
            paragraph = docAns.add_paragraph(text, style='List Number')
            # Format Options
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(8)


docQues.save('KahootExport-Question.docx')
print('\nSaved: KahootExport-Question.docx')
docAns.save('KahootExport-Answer.docx')
print('\nSaved: KahootExport-Answer.docx')

os.startfile('KahootExport-Question.docx')
os.startfile('KahootExport-Answer.docx')

input('\nHit ENTER to Exit\n')
