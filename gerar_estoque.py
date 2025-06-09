import pandas as pd
import random
from datetime import datetime, timedelta
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, alignment, Border, Side

# 1. Definindo Listas para geração de Dados Aleatórios
nome_base_livros = [
    "A Sombra do Vento", "1984", "Dom Quixote", "Orgulho e Preconceito", "Cem Anos de Solidão", "O Pequeno Príncipe", "Crime e Castigo", "O Senhor dos Anéis", "Harry Potter e a Pedra Filosofal", "O Alquimista", "A Revolução dos Bichos", "O Grande Gatsby", "Anna Karenina", "Ensaio sobre a Cegueira", "Memórias Póstumas de Brás Cubas", "Vinte Mil Léguas Submarinas"
]

editoras = [
    "Companhia das Letras", "Editora Rocco", "Intrínseca", "Editora Aleph", "HarperCollins Brasil", "Record", "Sextante", "Planeta",  "Globo Livros", "Alta Books"
]

# 2. Gerando os Dados do Estoque
num_livros_simulados = 50
estoque_livros = []

print("Gerando dados de estoque de livros...")

for i in range(num_livros_simulados):
    nome_livro = random.choice(nome_base_livros) + f" - vol. {random.randint(1, 3)}" if random.random() < 0.3 else random.choice(nome_base_livros)
    editora = random.choice(editoras)
    valor_custo = round(random.uniform(15.00, 80.00), 2)
    quantidade = random.randint(1, 100)
    notas = random.choice(["Novo", "Edição Especial", "Capa Dura", "Brochura", "Em Promoção", "Best-seller", ""])

    estoque_livros.append({
        "Nome do Livro": nome_livro,
        "Editora": editora,
        "Valor do Livro (Custo)": valor_custo,
        "Notas": notas,
        "Quantidade": quantidade
    })

print(f"Dados de {num_livros_simulados} livros gerados com sucesso!")

# 3. Criando um DataFrame Pandas
df_estoque = pd.DataFrame(estoque_livros)

# Exibindo as primeiras linhas do DataFrame para verificar
print("\nPrimeiras 5 linahs do DataFrame gerado:")
print(df_estoque.head())