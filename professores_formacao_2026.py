'''
Contabilizar todos os professores da rede para os Anos Finais do Ensino Fundamental e Ensino Médio. 
Essa informação virá do Relatório de Frequência Mensal de Professsores.
(SigEduc > Diário de Classe > Relatórios > Frequência > Relatório de Acompanhamento de Frequência Mensal Professor)

Vou ter lista nominal de todos os professores.
Como o mesmo professor pode dar aula em mais de uma escola e série, considerar a série e escola que ele tem maior número de aulas dadas.

Para Componentes Curriculares, considerar somente:
- "Língua Portuguesa",
- "Matemática".

Para Séries, considerar:
- 6º ano,
- 7º ano,
- 8º ano,
- 9º ano, 
- 1ª série,
- 2ª série,
- 3ª série.

Considerar somente Ensino Regular (excluir EJA e EPT).

Para identificação de escolas com ensino noturno, considerar o código das turmas.

'''
# Importação das bibliotecas
import pandas as pd
import glob
import os
from tqdm import tqdm  # Para barra de progresso
import numpy as np
import warnings
warnings.filterwarnings('ignore')
import openpyxl
import re


# RELATÓRIO DE FREQUÊNCIA de PROFESSORES- 2º TRIMESTRE (ABRIL, MAIO E JUNHO) - 2026
# caminho da pasta onde estão os arquivos
pasta = r"C:\Users\hugob\Downloads\Frequencia - Professores\2tri 2026"

# lista todos os arquivos .xlsx da pasta
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

# lista para armazenar os dataframes
dfs = []

for arquivo in tqdm(arquivos, desc="Processando arquivos"):
    # lê cada arquivo, pulando as 4 primeiras linhas
    df_unico = pd.read_excel(arquivo, skiprows=4)
    dfs.append(df_unico)

# concatena todos em um único dataframe
df_freq_prof_26 = pd.concat(dfs, ignore_index=True)


# DATAFRAMES RESERVAS
df_freq_prof_26_reserva = df_freq_prof_26.copy(deep=True)































