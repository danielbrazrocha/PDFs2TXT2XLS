#Converting multiple PDF files in the root folder to JPG (OCR/Pytesseract only read images)
import os
from pdf2image import convert_from_path

pdf_dir = r"./"
os.chdir(pdf_dir)

for pdf_file in os.listdir(pdf_dir):
    if pdf_file.endswith(".pdf"):
        pages = convert_from_path(pdf_file, 300)
        pdf_file = pdf_file[:-4]

        for page in pages:
            page.save("%s-page%d.jpg" % (pdf_file,pages.index(page)), "JPEG")

''' Skipping this...
#Verificando se a imagem está na pasta
from os import listdir
listdir('.')
'''

#Creating a text list of JPG files in the root folder
with open('listimg.txt', 'w') as arquivo:
    
    for jpg_file in os.listdir(pdf_dir):
        if jpg_file.endswith(".jpg"):
            arquivo.write(f"{jpg_file}\n")

#Converting JPG files in a unique TXT
import pytesseract
text = pytesseract.image_to_string('listimg.txt', lang='por')
'''print(text)'''


# Saving matches to a TXT
with open('texto.txt', 'w') as document:
    document.write(text)


#Using REGEX to find expressions (dd/yyyy nnn,nn OR dd/yyyy n.nnn,nn)
import re
regex = r"\d\d\D\d\d\d\d\s\d\d\d\D\d\d|\d\d\D\d\d\d\d\s\d\D\d\d\d\D\d\d"

#Saving the REGEX matches as a variable
matches = re.findall(regex, text)
'''print(matches)'''

# Converting the list to a tuple, splitting the content in two cols
res2col = [tuple(x.split(' ')) for x in matches]

#Convert the texto format to a Excel sheet
from datetime import datetime
import xlsxwriter
# Create a new workshhet
novo_arquivo = xlsxwriter.Workbook('Resultado.xlsx')
nova_planilha = novo_arquivo.add_worksheet()
bold = novo_arquivo.add_format({'bold': 1})
# Adding the values in the row 0 e na coluna 0
nova_planilha.write('A1', 'Mes/Ano', bold)
nova_planilha.write('B1', 'Valor', bold)
row = 1
col = 0
#Iterate data column by column
for mesano, valor in (res2col):
    nova_planilha.write(row, col,     mesano)
    nova_planilha.write(row, col + 1, valor)
    row += 1

nova_planilha.write(row, 0, 'Total')
nova_planilha.write(row, 1, 'Total2')

novo_arquivo.close()



