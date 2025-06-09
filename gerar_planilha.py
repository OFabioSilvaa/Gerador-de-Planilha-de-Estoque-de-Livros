import random
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

#Configurações do Projeto
NUM_LIVROS = 50
MARGEM_LUCRO = 0.30

# Listas de Dados
nomes_livros_base = [
    "A Arte da Guerra", "1984", "Dom Quixote", "O Pequeno Príncipe",
    "Cem Anos de Solidão", "Crime e Castigo", "Orgulho e Preconceito",
    "O Senhor dos Anéis", "Harry Potter e a Pedra Filosofal", "O Alquimista",
    "A Metamorfose", "Ulisses", "Moby Dick", "As Crônicas de Nárnia",
    "O Guia do Mochileiro das Galáxias", "O Apanhador no Campo de Centeio",
    "Em Busca do Tempo Perdido", "Grande Sertão: Veredas", "Memórias Póstumas de Brás Cubas",
    "Vidas Secas"
]

editoras = [
    "Editora Alfa", "Editora Beta", "Editora Gama", "Editora Delta",
    "HarperCollins", "Penguin Random House", "Rocco", "Companhia das Letras"
]

def gerar_dados_livro():
    # Gerando dados fictícios para um único livro.
    nome = random.choice(nomes_livros_base) + f" - Vol. {random.randint(1, 3)}" if random.random() < 0.3 else random.choice(nomes_livros_base)
    editora = random.choice(editoras)
    valor_compra = round(random.uniform(15.00, 150.00), 2) 
    notas = random.randint(1, 5) 
    quantidade = random.randint(1, 200) 

    return {
        "Nome do Livro": nome,
        "Editora": editora,
        "Valor de Compra": valor_compra,
        "Notas": notas,
        "Quantidade": quantidade
    }

def criar_planilha_estoque(nome_arquivo="estoque_livros.xlsx"):
    # Cria e preenche a planilha Excel com dados de estoque.
    wb = Workbook()
    ws = wb.active
    ws.title = "Estoque de Livros"
    # Cabeçalhos 
    headers = [
        "Nome do Livro", "Editora", "Valor de Compra", "Notas", "Quantidade",
        "Preço de Venda", "Lucro Total", "Valor Total"
    ]
    ws.append(headers)
    # Geração e Inserção de Dados
    for i in range(NUM_LIVROS):
        dados_livro = gerar_dados_livro()
        linha_dados = [
            dados_livro["Nome do Livro"],
            dados_livro["Editora"],
            dados_livro["Valor de Compra"],
            dados_livro["Notas"],
            dados_livro["Quantidade"]
        ]
        ws.append(linha_dados)

        current_row = ws.max_row
        # Colunas:
        # C: Valor de Compra
        # E: Quantidade
        # F: Preço de Venda
        # G: Lucro Total
        # H: Valor Total

        # Fórmula para Preço de Venda (Coluna F): Valor de Compra * (1 + MARGEM_LUCRO)
        ws[f'F{current_row}'] = f'={get_column_letter(3)}{current_row}*(1+{MARGEM_LUCRO})'
        # Fórmula para Lucro Total (Coluna G): (Preço de Venda - Valor de Compra) * Quantidade
        ws[f'G{current_row}'] = f'=({get_column_letter(6)}{current_row}-{get_column_letter(3)}{current_row})*{get_column_letter(5)}{current_row}'
         # Fórmula para Valor Total (Coluna H): Preço de Venda * Quantidade
        ws[f'H{current_row}'] = f'={get_column_letter(6)}{current_row}*{get_column_letter(5)}{current_row}'
        
    for col in ws.columns:
        max_length = 0
        column = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    # Salvar a planilha
    wb.save(nome_arquivo)
    print(f"Planilha '{nome_arquivo}' criada com sucesso!")

if __name__ == "__main__":
    criar_planilha_estoque()
