# Percorrendo pasta e arquivos
import os
import docx2txt

def escreverArquivo(nome: str, count: int):
  with open('paginasContadas.txt', 'r') as arquivo:
    conteudo = arquivo.readlines()
    conteudo.append(f'Nome do arquivo: {nome}')
    conteudo.append(f'\nQtd Caracteres: {count}\n\n')

  with open('paginasContadas.txt', 'w') as arquivo:
    arquivo.writelines(conteudo)

def escreverTotal(countT: int):
  with open('paginasContadas.txt', 'r') as arquivo:
    conteudo = arquivo.readlines()
    conteudo.append(f'========================\n')
    conteudo.append(f'Total de caracteres: {countT}')

  with open('paginasContadas.txt', 'w') as arquivo:
    arquivo.writelines(conteudo)

# variaveis
files = []
path = f'{os.path.dirname(os.path.realpath(__file__))}'
countTotal = 0

# percorrer, add arquivos na lista
for (dirpath, dirnames, filenames) in os.walk(path):
  files.extend(filenames)
  break

# criar arquivo vazio
arquivo = open('paginasContadas.txt', 'w')
arquivo.close()

#percorrer lista e add no txt nome e contagem
for filename in files:
  if str(filename).__contains__('.docx'):
    #print(filename)

    word = docx2txt.process(f'{filename}')
    #print(filename)
    count = (len(word)-word.count('\n'))
    countTotal += count
    escreverArquivo(filename, count)
escreverTotal(countTotal)

#print(files)