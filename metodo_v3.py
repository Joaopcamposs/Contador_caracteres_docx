import os
import re
import zipfile
from bibliotecas.docx2pdf_file import convert
from bibliotecas.docx2txt_file import xml2text
import pdfplumber


## Variaveis
filename = 'FIBO_PV_FIS-1_C1_CE_cap6_T_A_T'


## Unzip the docx in memory
def unzipFiles(filename: str):
    zipf = zipfile.ZipFile(f'{filename}.docx')
    filelist = zipf.namelist()
    return zipf, filelist


def formatarTexto(text: str):
    while text.__contains__('  '):
        text = text.replace('\n', '')
        text = text.replace('  ', ' ')
        text = text.strip()
        #print(text)
    return text


## Contar caracteres do cabecalho
def contarCabecalho(filename: str):
    text = ''
    zipf, filelist = unzipFiles(filename)
    header_xmls = 'word/header[0-9]*.xml'
    for fname in filelist:
        if re.match(header_xmls, fname):
            text += xml2text(zipf.read(fname))
    #print(text)
    #print(len(text)-text.count('\n'))
    countCabecalho = len(text)-text.count('\n')
    return countCabecalho


## Criar arquivo pdf
def criarPDF(filename: str):
    convert(f'{filename}.docx')


## Ler Todas as paginas do arquivo pdf e contar
def contarPDF(filename: str):
    with pdfplumber.open(f'{filename}.pdf') as pdf:
        #print(len(pdf.pages))
        i=0
        qtdPaginas = len(pdf.pages)
        contChar = 0
        countCabecalho = contarCabecalho(filename)
        while(i < qtdPaginas):
            pagina = pdf.pages[i]
            text = pagina.extract_text()
            text = formatarTexto(text)
            #print(text)
            contChar += (len(text))
            #print(len(text)-text.count('\n'))
            i+=1
        print(f'Qtd Caracteres = {contChar - (countCabecalho * qtdPaginas)}')
        return contChar - (countCabecalho * qtdPaginas)


## Gerar arquivo em txt com todo o texto do pdf para possivel conferencia
def criarTXT(filename: str):
    with pdfplumber.open(f'{filename}.pdf') as pdf:
        ## Criar arquivo vazio
        arquivo = open(f'{filename}.txt', 'w', encoding="utf-8")
        arquivo.close()

        i=0
        qtdPaginas = len(pdf.pages)
        while i < qtdPaginas:
            with open(f'{filename}.txt', 'r', encoding="utf-8") as arquivo:
                conteudo = arquivo.readlines()
                #print(pdf.pages[i].extract_text())
                texto = formatarTexto(pdf.pages[i].extract_text())
                #texto = pdf.pages[i].extract_text()
                conteudo.append(texto)
                #print(conteudo)
            with open(f'{filename}.txt', 'w', encoding="utf-8") as arquivo:
                arquivo.writelines(conteudo)
            i+=1

## Excluir arquivos
def excluirArquivos(filename: str):
    os.remove(f'{filename}.txt')
    os.remove(f'{filename}.pdf')


## Main
criarPDF(filename)
cont = contarPDF(filename)
criarTXT(filename)
#excluirArquivos(filename)

