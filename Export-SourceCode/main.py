import string
import requests
from docx import Document

answerURL = 'https://create.kahoot.it/rest/kahoots/{}/card/?includeKahoot=true'.format(input().split('/')[-1])
data = requests.get(answerURL).json()['kahoot']['questions']
doc = Document()


for i in range(0, len(data)):
    paragraph = doc.add_paragraph(
        '{quesNum}. {ques}'.format(quesNum=i+1, ques=data[i]['question']))
    paragraph.paragraph_format.line_spacing = 1.0
    # print('{quesNum}. {ques}'.format(quesNum=i+1, ques=data[i]['question']))
    for j in range(0, len(data[i]['choices'])):
        paragraph = doc.add_paragraph(
            '\t{ansNum}. {ans}'.format(ansNum=list(string.ascii_uppercase)[j], ans=data[i]['choices'][j]['answer']))
        paragraph.paragraph_format.line_spacing = 1.0
        # print('\t{ansNum}. {ans}'.format(ansNum=list(string.ascii_uppercase)[j], ans=data[i]['choices'][j]['answer']))

doc.save('KahootCapture.docx')
