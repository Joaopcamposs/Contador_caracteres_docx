import os
import docx2txt

## Funcao para escrever nome do arquivo e qtd de paginas no txt
def escreverArquivo(nome: str, count: int):
  with open('paginasContadas2.txt', 'r', encoding="utf-8") as arquivo:
    conteudo = arquivo.readlines()
    conteudo.append(f'Nome do arquivo: {nome}')
    conteudo.append(f'\nQtd Caracteres: {count}\n\n')

  with open('paginasContadas2.txt', 'w', encoding="utf-8") as arquivo:
    arquivo.writelines(conteudo)

## Funcao para escrever o total no final do arquivo
def escreverTotal(countT: int):
  with open('paginasContadas2.txt', 'r', encoding="utf-8") as arquivo:
    conteudo = arquivo.readlines()
    conteudo.append(f'========================\n')
    conteudo.append(f'Total de caracteres: {countT}')

  with open('paginasContadas2.txt', 'w', encoding="utf-8") as arquivo:
    arquivo.writelines(conteudo)

## Funcao que remove duplicados, escrevendo em um txt e retorna contagem de caracteres
def RemoverDuplicados(filename: str):
  with open('cont1.txt', 'w', encoding="utf-8") as cont1:
    word = docx2txt.process(f'{filename}')
    cont1.writelines(word)

  lines_seen = set()  # holds lines already seen
  outfile = open('cont2.txt', "w", encoding="utf-8")
  for line in open('cont1.txt', "r", encoding="utf-8"):
    if line not in lines_seen:  # not a duplicate
      outfile.write(line)
      lines_seen.add(line)
  outfile.close()

  #contagem de caracteres
  with open('cont2.txt', 'r', encoding="utf-8") as arquivo2:
    word2 = arquivo2.read()
    count2 = (len(word2) - word2.count('\n'))

  os.remove('cont1.txt')
  os.remove('cont2.txt')

  return count2

# variaveis
files = []
path = f'{os.path.dirname(os.path.realpath(__file__))}'
countTotal = 0

# percorrer, add arquivos na lista
for (dirpath, dirnames, filenames) in os.walk(path):
  files.extend(filenames)
  break

# criar arquivo vazio
arquivo = open('paginasContadas2.txt', 'w', encoding="utf-8")
arquivo.close()

#percorrer lista e add no txt nome e contagem
for filename in files:
  if str(filename).__contains__('.docx'):
    #print(filename)
    #print(filename)
    count = RemoverDuplicados(filename)
    #print('passou')
    countTotal += count
    escreverArquivo(filename, count)
escreverTotal(countTotal)

#print(files)