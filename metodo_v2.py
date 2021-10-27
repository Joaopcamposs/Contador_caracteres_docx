import os
from bibliotecas import docx2txt_file
import zipfile


## Separar via zip e extrair texto
zipf = zipfile.ZipFile('caixa_texto.docx')
filelist = zipf.namelist()
#print(filelist)

doc_xml = 'word/document.xml'
text = ''
text += docx2txt_file.xml2text(zipf.read(doc_xml))
#text = text.replace('\n', '')
with open('teste.txt', 'w') as arquivo:
    arquivo.writelines(text)
#print(text)

## Variaveis para contar
cont1 = 0
cont2 = 0

## Remover duplicados e escrever no arquivo novo
lines_seen = set() # holds lines already seen
outfile = open('teste2.txt', "w")
for line in open('teste.txt', "r"):
    if line not in lines_seen: # not a duplicate
        outfile.write(line)
        lines_seen.add(line)
outfile.close()

## Contar total de caracteres
with open('teste.txt', 'r') as arquivo:
    word = arquivo.read()
    count1 = (len(word) - word.count('\n'))

with open('teste2.txt', 'r') as arquivo2:
    word2 = arquivo2.read()
    count2 = (len(word2) - word2.count('\n'))

os.remove('teste.txt')
os.remove('teste2.txt')

print(f'Contagem: {count1}')
print(f'Contagem sem duplicados: {count2}')
