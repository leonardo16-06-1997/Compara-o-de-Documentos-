import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
from difflib import SequenceMatcher
from tkinter import ttk

# variáveis globais
arquivo_padrao = ''
arquivo_aluno = ''

# função para extrair texto de um arquivo Word
def extrair_texto_word(arquivo):
    doc = docx.Document(arquivo)
    texto = ''
    for paragraph in doc.paragraphs:
        texto += paragraph.text + '\n'
    return texto

# função que é chamada quando o botão de comparação é pressionado
def comparar():
    try:
        # determinar se os arquivos são PDF ou Word
        pdf_extensoes = ('.pdf', '.PDF')
        word_extensoes = ('.docx', '.DOCX')

        if arquivo_padrao.endswith(pdf_extensoes) and arquivo_aluno.endswith(pdf_extensoes):
            with open(arquivo_padrao, 'rb') as file_padrao, open(arquivo_aluno, 'rb') as file_aluno:
                leitor_padrao = PyPDF2.PdfReader(file_padrao)
                texto_padrao = ''
                for pagina in leitor_padrao.pages:
                    texto_padrao += pagina.extract_text()

                leitor_aluno = PyPDF2.PdfReader(file_aluno)
                texto_aluno = ''
                for pagina in leitor_aluno.pages:
                    texto_aluno += pagina.extract_text()
        elif arquivo_padrao.endswith(word_extensoes) and arquivo_aluno.endswith(word_extensoes):
            texto_padrao = extrair_texto_word(arquivo_padrao)
            texto_aluno = extrair_texto_word(arquivo_aluno)
        else:
            raise ValueError('Os arquivos devem ser ambos PDF ou Word (docx)')

        # calcular o grau de semelhança entre os textos
        semelhanca = SequenceMatcher(None, texto_padrao, texto_aluno).ratio()
        percentual = semelhanca * 100

        # atualizar a label com o resultado
        resultado_label.config(text=f'O percentual de semelhança é {percentual:.2f}%', font=('Times New Roman', 16))
    except Exception as e:
        resultado_label.config(text=f'Ocorreu um erro ao comparar os arquivos: {e}')

# função que é chamada quando o botão de upload do arquivo padrão é pressionado
def selecionar_padrao():
    global arquivo_padrao
    arquivo_padrao = filedialog.askopenfilename(filetypes=[("Arquivos suportados", "*.pdf;*.docx;*.PDF;*.DOCX")])
    if arquivo_padrao != '':
        padrao_button.config(text=f'Arquivo Padrão: {arquivo_padrao}')

# função que é chamada quando o botão de upload do arquivo do aluno é pressionado
def selecionar_aluno():
    global arquivo_aluno
    arquivo_aluno = filedialog.askopenfilename(filetypes=[("Arquivos suportados", "*.pdf;*.docx;*.PDF;*.DOCX")])
    if arquivo_aluno != '':
        aluno_button.config(text=f'Arquivo do Aluno: {arquivo_aluno}')

# criar a janela
janela = tk.Tk()
janela.title('Comparação de Documentos')

# definir o tamanho da janela
janela.geometry('1000x600')

# carregar a imagem de fundo
fundo = Image.open('fundo1.png')
fundo = fundo.resize((1000, 600), Image.ANTIALIAS)
fundo = ImageTk.PhotoImage(fundo)

# criar um canvas para exibir a imagem de fundo
canvas = tk.Canvas(janela, width=1000, height=600)
canvas.pack(fill='both', expand=True)
canvas.create_image(0, 0, image=fundo, anchor='nw')

# Criar o cabeçalho do programa com instruções
rodape = tk.Label(janela, text="ANALISADOR DE SIMILARIDADE DE DOCUMENTOS", font=("Times New Roman", 16, 'bold'), bg='white', wraplength=750, justify="center")
canvas.create_window(500, 50, window=rodape)

style = ttk.Style()
style.configure('Custom.TButton',
                font=('Times New Roman', 16),
                background='black',
                foreground='black',
                relief='raised',
                padding=(10, 10))

padrao_button = ttk.Button(janela, text='Selecionar Arquivo de Referência', command=selecionar_padrao, style='Custom.TButton')
canvas.create_window(500, 200, window=padrao_button)

aluno_button = ttk.Button(janela, text='Selecionar Arquivo de Comparação', command=selecionar_aluno, style='Custom.TButton')
canvas.create_window(500, 300, window=aluno_button)

comparar_button = ttk.Button(janela, text='Comparar', command=comparar, style='Custom.TButton')
canvas.create_window(500, 400, window=comparar_button)

# criar a label para exibir o resultado
resultado_label = tk.Label(janela, text='', bg='white')
canvas.create_window(500, 500, window=resultado_label)

# criar o rodapé do programa com instruções
rodape = tk.Label(janela, text="Selecione os arquivos (PDF ou Word) para comparar e clique em 'Comparar' para verificar a similaridade.", font=("Arial", 12), bg='white', wraplength=750, justify="center")
canvas.create_window(500, 550, window=rodape)

# iniciar a janela
janela.mainloop()