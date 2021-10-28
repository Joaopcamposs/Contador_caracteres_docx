import os
import re
import zipfile
from bibliotecas.docx2pdf_file import convert
from bibliotecas.docx2txt_file import xml2text
import pdfplumber

## Variaveis
files = []
path = f'{os.path.dirname(os.path.realpath(__file__))}'
contTotal = 0

## Unzip the docx in memory
def unzipFiles(filename: str):
    zipf = zipfile.ZipFile(f'{filename}')
    filelist = zipf.namelist()
    return zipf, filelist

def formatarTexto(text: str):
    while text.__contains__('  '):
        text = text.replace('\n', '')
        text = text.replace('  ', ' ')
        text = text.strip()
        # print(text)
    return text

## Funcao para escrever nome do arquivo e qtd de paginas no txt
def escreverArquivo(nome: str, count: int):
  with open('paginasContadas3.txt', 'r', encoding="utf-8") as arquivo:
    conteudo = arquivo.readlines()
    conteudo.append(f'Nome do arquivo: {nome}')
    conteudo.append(f'\nQtd Caracteres: {count}\n\n')

  with open('paginasContadas3.txt', 'w', encoding="utf-8") as arquivo:
    arquivo.writelines(conteudo)

## Funcao para escrever o total no final do arquivo
def escreverTotal(countT: int):
  with open('paginasContadas3.txt', 'r', encoding="utf-8") as arquivo:
    conteudo = arquivo.readlines()
    conteudo.append(f'========================\n')
    conteudo.append(f'Total de caracteres: {countT}')

  with open('paginasContadas3.txt', 'w', encoding="utf-8") as arquivo:
    arquivo.writelines(conteudo)

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
    convert(f'{filename}')

## Ler Todas as paginas do arquivo pdf e contar
def contarPDF(filename: str):
    global contTotal
    nomePDF = filename.replace('.docx', '.pdf')
    with pdfplumber.open(f'{nomePDF}') as pdf:
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
        contTotal += contChar-(countCabecalho*qtdPaginas)
        return contChar-(countCabecalho*qtdPaginas)

## Gerar arquivo em txt com todo o texto do pdf para possivel conferencia
def gerarTXT(filename: str):
    nomeTXT = filename.replace('.docx', '.txt')
    nomePDF = filename.replace('.docx', '.pdf')
    with pdfplumber.open(f'{nomePDF}') as pdf:
        ## Criar arquivo vazio
        arquivo = open(f'{nomeTXT}', 'w', encoding="utf-8")
        arquivo.close()

        i=0
        qtdPaginas = len(pdf.pages)
        while i < qtdPaginas:
            with open(f'{nomeTXT}', 'r', encoding="utf-8") as arquivo:
                conteudo = arquivo.readlines()
                #print(pdf.pages[i].extract_text())
                texto = formatarTexto(pdf.pages[i].extract_text())
                conteudo.append(texto)
                #print(conteudo)

            with open(f'{nomeTXT}', 'w', encoding="utf-8") as arquivo:
                arquivo.writelines(conteudo)

            i+=1

def excluirArquivos(filename: str):
    nomeTXT = filename.replace('.docx', '.txt')
    nomePDF = filename.replace('.docx', '.pdf')
    #os.remove(f'{nomeTXT}')
    os.remove(f'{nomePDF}')

## Criar arquivo vazio
arquivo = open('paginasContadas3.txt', 'w', encoding="utf-8")
arquivo.close()

## Percorrer, add arquivos na lista
for (dirpath, dirnames, filenames) in os.walk(path):
  files.extend(filenames)
  break

## Percorrer lista e add no txt nome e contagem
for filename in files:
  if str(filename).__contains__('.docx'):
    #print(filename)
    criarPDF(filename)
    cont = contarPDF(filename)
    escreverArquivo(filename, cont)
    excluirArquivos(filename)

escreverTotal(contTotal)
