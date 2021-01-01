import string
import requests
from docx import Document
from docx.shared import Pt

answerURL = 'https://create.kahoot.it/rest/kahoots/{}/card/?includeKahoot=true'.format(input().split('/')[-1])
data = requests.get(answerURL).json()['kahoot']['questions']
docQues = Document()


for i in range(0, len(data)):
    # Paragraph of Questions
    paragraph = docQues.add_paragraph(
        '{ques}'.format(ques=data[i]['question']), style='List Number')
    # Format Options
    paragraph.paragraph_format.line_spacing = 1.0
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(8)
    # print('{quesNum}. {ques}'.format(quesNum=i+1, ques=data[i]['question']))
    for j in range(0, len(data[i]['choices'])):
        # Paragraph of Choices
        paragraph = docQues.add_paragraph(
            '\t{ansNum}. {ans}'.format(ansNum=list(string.ascii_uppercase)[j], ans=data[i]['choices'][j]['answer']))
        # Format Options
        paragraph.paragraph_format.line_spacing = 1.0
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(8)
        # print('\t{ansNum}. {ans}'.format(ansNum=list(string.ascii_uppercase)[j], ans=data[i]['choices'][j]['answer']))

docQues.save('KahootExport-Question.docx')
