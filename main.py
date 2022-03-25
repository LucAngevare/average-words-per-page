import json, re, os, subprocess as subp, time, openpyxl
from datetime import datetime
from docx import Document
from PyPDF2 import PdfFileReader

dictionary = json.load(open("index.json"))

def count_chars(doc):
    newparatextlist = []
    for paratext in doc.paragraphs:
        newparatextlist.append(paratext.text)
    
    return len(re.findall(r'\w+', '\n'.join(newparatextlist)))

for dictionary_iteration in range(10000): #going for 1.000 iterations here just to see exactly how long this takes, if all goes well and stuff I'll up it, calculate how long each iteration takes exactly and I guess run the simulations overnight
    for word_amount in range(300, 700): #using a starting point of 300 words because on average 500 words fills a page so I doubt any less will do, it's always possible though, a maximum of 700 because why not, it'll start the next dictionary_iteration when it has 2 pages
        doc = Document()
        wb = openpyxl.load_workbook("data.xlsx")
        worksheet = wb["Sheet1"]
        startingTime = datetime.now()
        text = " ".join(dictionary[dictionary_iteration:dictionary_iteration+word_amount])
        paragraph = doc.add_paragraph(text)
        doc.save("sim.docx")
        subp.run('pandoc -o sim.pdf sim.docx')
        with open('sim.pdf', 'rb') as f:
            pdf = PdfFileReader(f)
            pageNum = pdf.getNumPages()
        maximum_row = worksheet.max_row+1
        worksheet[f'A{maximum_row}'] = dictionary_iteration
        if (count_chars(doc)==word_amount): worksheet[f'B{maximum_row}'] = count_chars(doc)
        else:
            print(f'{count_chars(doc)} not equal to {word_amount}!')
            worksheet[f'B{maximum_row}'] = (count_chars(doc)+word_amount)/2 #if something does go wrong and the actual word count in the document that is read in within the data is not equal to the theoretical word count, it should just take the average as that probably won't have too much impact on the simulation we're running, I will log this to the console, so if it happens constantly, something consistently doesn't add up, and more research is required.
        worksheet[f'C{maximum_row}'] = pageNum
        worksheet[f'D{maximum_row}'] = datetime.now().microsecond-startingTime.microsecond
        worksheet[f'E{maximum_row}'] = len(text)
        print(f'{dictionary_iteration} with a length of {word_amount} has {pageNum} pages, and {len(text)} characters.')
        result = False
        os.remove("sim.docx")
        while not result:
            try:
                os.remove("sim.pdf")
                result = True
            except:
                time.sleep(1)
        wb.save("data.xlsx")
        if (pageNum == 2):
            break
    else:
        continue