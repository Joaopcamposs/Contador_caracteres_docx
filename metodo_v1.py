import docx2txt

# TXT ===============================================================
#adicionando encoding, nao conta acento como 2 caracteres
with open('txt2.txt', 'r', encoding='utf-8') as arquivo:
    txt = arquivo.read()

print('===== TXT ======')
#print(f'{txt}')
print(f'tamanho: {len(txt)}')
print('---- Quebra de linha conta +1 ----')
cont = txt.count('\n')
print(f'Quebras de linha: {cont}')
print(f'TOTAL = {len(txt)-cont}')

# DOCX =============================================================
word = docx2txt.process("artigo.docx")

print('\n===== DOCX ======')
#print(f'{word}')
print(f'tamanho: {len(word)}')
print('---- Quebra de linha conta +2 ----')
contQuebras = word.count('\n')
print(f'Quebras de linha: {contQuebras}')
cont = (len(word)-word.count('\n'))
print(f'TOTAL = {cont}')
