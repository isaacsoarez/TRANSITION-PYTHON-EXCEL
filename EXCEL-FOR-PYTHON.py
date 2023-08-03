import openpyxl

# 1° PASSO (CARREGAR ARQUIVO)
book = openpyxl.load_workbook('DADOS.xlsx')

# 2° PASSO (MOSTRAR PLANILHAS EXISTENTES)
print(book.sheetnames)

# 3° PASSO (CRIAR NOVA PLANILHA)
book.create_sheet('DADOS02')
print(book.sheetnames)

# 4° PASSO (COLOCAR NOVOS CONTEÚDOS NA PLANILHA DADOS02)
DADOS02_page = book["DADOS02"]
DADOS02_page.append(["NOME", "IDADE", "SÉRIE"])
DADOS02_page.append(["Maria", "19", "3°ano"])
DADOS02_page.append(["Pedro", "18", "2°ano"])
DADOS02_page.append(["Lucas", "21", "5°ano"])

# 5° PASSO (MOSTRAR O CONTEÚDO DA PLANILHA DADOS01)
DADOS02_page = book["DADOS02"]
for rows in DADOS02_page.iter_rows(min_row=1, max_row=7):
    print(rows[0].value, rows[1].value, rows[2].value)

# 6° PASSO (SALVAR O ARQUIVO)
book.save('DADOS.xlsx')
