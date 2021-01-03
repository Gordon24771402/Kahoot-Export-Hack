import requests
from docx import Document
from docx.shared import Pt
import string
import platform
import subprocess
import os

# Create Word Document for Questions and Answers
docQues = Document()
docAns = Document()

# List Copyright Info and Instruction to Continue
print('Copyright (c) 2020-2021 Hao Kang')
print('Hit ENTER to Continue Export.\n')

# Loop-1: Read Multiple URLs
while True:

    # Read URL
    URL = input('Copy and Paste Kahoot URL: ')

    # Whether to Continue or Break
    if URL != '':
        kahootID = URL.split('/')[-1]
        answerURL = 'https://create.kahoot.it/rest/kahoots/{}/card/?includeKahoot=true'.format(kahootID)
        data = requests.get(answerURL).json()['kahoot']['questions']

        # Loop-1-1: Questions
        for i in range(0, len(data)):
            text = '{ques}'.format(ques=data[i]['question'])
            # Remove &nbsp;
            text = text.replace('&nbsp;', '')
            # Remove <i> and </i>
            text = text.replace('<i>', '')
            text = text.replace('</i>', '')
            # Remove <b> and </b>
            text = text.replace('<b>', '')
            text = text.replace('</b>', '')
            # Paragraph of Questions
            paragraph = docQues.add_paragraph(text, style='List Number')
            # Format Options
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(8)

            # Loop-1-1-1: Choices
            for j in range(0, len(data[i]['choices'])):
                text = '\t{ansNum}. {ans}'.format(ansNum=list(string.ascii_uppercase)[j], ans=data[i]['choices'][j]['answer'])
                # Remove &nbsp;
                text = text.replace('&nbsp;', '')
                # Remove <i> and </i>
                text = text.replace('<i>', '')
                text = text.replace('</i>', '')
                # Remove <b> and </b>
                text = text.replace('<b>', '')
                text = text.replace('</b>', '')
                # Paragraph of Choices
                paragraph = docQues.add_paragraph(text)
                # Format Options
                paragraph.paragraph_format.line_spacing = 1.0
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(8)

            # Loop-1-1-2: Answers
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
    else:
        break

# Save Documents and Notify Each
docQues.save('KahootExport-Question.docx')
print('\nSaved: KahootExport-Question.docx')
docAns.save('KahootExport-Answer.docx')
print('Saved: KahootExport-Answer.docx')

# User Experience: Keep Program Run and Open Documents if Necessary
input('\n(* Hit ENTER to Exit and Open Documents)\n')

# Mac
if platform.system() == 'Darwin':
    subprocess.call(('open', 'KahootExport-Question.docx'))
    subprocess.call(('open', 'KahootExport-Answer.docx'))
# Windows
elif platform.system() == 'Windows':
    os.startfile('KahootExport-Question.docx')
    os.startfile('KahootExport-Answer.docx')
# Linux
else:
    subprocess.call(('xdg-open', 'KahootExport-Question.docx'))
    subprocess.call(('xdg-open', 'KahootExport-Answer.docx'))
