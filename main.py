
from module.process_pdf import process_pdf
from module.structuring_exel import structuring


# Solicitar informações ao usuário
pdf_path = input("Informe o caminho do PDF: ")
pages = str(input("Informe as páginas a serem processadas(ex: 1, 1-3, all): "))
name_file = str(input("Informe o Nome de arquivo que deseja: "))

# Chamar a função e exibir o resultado
result = process_pdf(pdf_path, pages)
process = structuring(name_file)

print(result)
print(process)