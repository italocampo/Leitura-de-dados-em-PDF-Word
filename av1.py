# Aluno: Francisco Ítalo de Souza Campos
# Turma: ADS 3º Semestre
# Matrícula: 2024012967
# Descrição: Script que lê números de arquivos .pdf ou .docx e calcula estatísticas.


#Primeiramente, aqui eu importo algumas bibliotecas.
import os #Serve para montar caminhos
import statistics #Usada para calcular médias e medianas.
from docx import Document #Permite que o código leia arquivos .docx
import fitz #Ler arquivos PDF (PyMuPDF)

pasta = 'documentos' #Definindo o lugar dos arquivos
docx_path = os.path.join(pasta,'dados.docx') #Crio caminhos
pdf_path = os.path.join(pasta, 'dados.pdf')

numeros = [] #Aqui é uma lista onde estará todos os números lidos

if os.path.exists(docx_path): #Aqui começa a leitura do documento, primeiramente se ele existir. Após a condição,o for vai percorrer todos os parágrafos e no final faz uma validação usando o Try except.
    print('Lendo arquivo DOCX...')
    doc = Document(docx_path)
    for paragrafo in doc.paragraphs:
        try:
            num = float(paragrafo.text.strip())
            numeros.append(num)
        except ValueError:
            continue


elif os.path.exists(pdf_path): #Do mesmo jeito à cima, aqui se repete com alguns detalhes diferentes: Criei dois For para o primeiro percorrer as páginas e o segundo converter cada linha em números.
    print('Lendo arquivo PDF...')
    doc = fitz.open(pdf_path)
    for page in doc:
        linhas = page.get_text().splitlines()
        for linha in linhas:
            try:
                num = float(linha.strip())
                numeros.append(num)
            except ValueError:
                continue

else:
    print("Nenhum arquivo 'dados.pdf' ou 'dados.docx' encontrado na pasta documentos.")
    exit()

if not numeros:
    print("Arquivo encontrado, porém não há números válidos. ")
    exit()

#Aqui saem os valores calculados.
print("\n📊 Estatísticas dos dados:")
print(f"Quantidade de números: {len(numeros)}")
print(f"Média: {statistics.mean(numeros):.2f}")
print(f"Mediana: {statistics.median(numeros):.2f}")
print(f"Somatório: {sum(numeros):.2f}")
print(f"Maior valor: {max(numeros)}")
print(f"Menor valor: {min(numeros)}")
