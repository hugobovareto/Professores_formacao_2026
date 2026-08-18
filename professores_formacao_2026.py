'''
Contabilizar todos os professores da rede para os Anos Finais do Ensino Fundamental e Ensino Médio. 
Essa informação vem do Relatório de Frequência Mensal de Professsores.
(SigEduc > Diário de Classe > Relatórios > Frequência > Relatório de Acompanhamento de Frequência Mensal Professor)

Vou ter lista nominal de todos os professores.
Como o mesmo professor pode dar aula em mais de uma escola e série, considerar a série e escola que ele tem maior número de aulas dadas.
Como a frequência de lançamento das aulas é baixa (em muitos casos, 'AULAS DADAS' = 0), vou considerar 'AULAS PREVISTAS' para definir o local do professor.

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


# Manter somente os componentes de interesse (Língua Portuguesa e Matemática)
componentes = ['Língua Portuguesa',
                'Matemática']

df_freq_prof_26 = df_freq_prof_26[df_freq_prof_26['COMPONENTE'].isin(componentes)]


# Manter somente as Séries de interesse (Anos Finais Ensino Fundamentla e Ensino Médio)
series = ['1ª SÉRIE',
          '2ª SÉRIE',
          '3ª SÉRIE',
          '6º ANO',
          '7º ANO',
          '8º ANO',
          '9º ANO']

df_freq_prof_26 = df_freq_prof_26[df_freq_prof_26['SÉRIE'].isin(series)]


# Manter somente as etapas de ensino de interesse (excluir EPT)
df_freq_prof_26 = df_freq_prof_26[df_freq_prof_26['ETAPA DE ENSINO'].isin(['ENSINO MÉDIO POTIGUAR', 'ENSINO FUNDAMENTAL'])]


# Eliminar duplicatas de professores (MATRÍCULA) mantendo somente o maior número de 'AULAS PREVISTAS'
df_freq_prof_26 = df_freq_prof_26.sort_values(by=['MATRÍCULA', 'AULAS PREVISTAS'], ascending=False)
df_freq_prof_26 = df_freq_prof_26.drop_duplicates(subset=['MATRÍCULA'], keep='first')


# São 1.713 vagas para os professores, incluindo noturno

len(df_freq_prof_26)
# Tem 1.913 professores na base, incluindo Noturno, então só 200 ficam de fora da Formação.


# Turmas noturnas
sorted(df_freq_prof_26['TURMA'].unique())
# Tem as seguintes turmas noturnas: 'EMPN1A', 'EMPN1B', 'EMPN1C', 'EMPN1D', 'EMPN2A', 'EMPN2B', 'EMPN2C', 'EMPN3A', 'EMPN3B', 'EMPN3C'


# Apenas os professores com turmas noturnas
df_noturno = df_freq_prof_26[df_freq_prof_26['TURMA'].isin(['EMPN1A', 'EMPN1B', 'EMPN1C', 'EMPN1D', 'EMPN2A', 'EMPN2B', 'EMPN2C', 'EMPN3A', 'EMPN3B', 'EMPN3C'])]

# Professores totais, menos do noturno
df_regular = df_freq_prof_26[~df_freq_prof_26['TURMA'].isin(['EMPN1A', 'EMPN1B', 'EMPN1C', 'EMPN1D', 'EMPN2A', 'EMPN2B', 'EMPN2C', 'EMPN3A', 'EMPN3B', 'EMPN3C'])]


# Agrupar os quantitativos de professores por escola
# REGULAR
# Criar indicadores para cada linha
df_regular['Professores EF-AF'] = df_regular['SÉRIE'].isin([
    '6º ANO', '7º ANO', '8º ANO', '9º ANO'
])

df_regular['Professores EM'] = df_regular['SÉRIE'].isin([
    '1ª SÉRIE', '2ª SÉRIE', '3ª SÉRIE'
])

# Agrupar por escola
df_regular_escola = (
    df_regular
    .groupby('INEP ESCOLA')
    .agg(
        DIREC=('DIREC', lambda x: x.mode().iloc[0]),
        MUNICÍPIO=('MUNICÍPIO', lambda x: x.mode().iloc[0]),
        ESCOLA=('ESCOLA', lambda x: x.mode().iloc[0]),
        **{
            'Professores EF-AF': ('Professores EF-AF', 'sum'),
            'Professores EM': ('Professores EM', 'sum')
        }
    )
    .reset_index()
)


# NOTURNO
# Criar indicadores para cada linha
df_noturno['Professores EF-AF'] = df_noturno['SÉRIE'].isin([
    '6º ANO', '7º ANO', '8º ANO', '9º ANO'
])

df_noturno['Professores EM'] = df_noturno['SÉRIE'].isin([
    '1ª SÉRIE', '2ª SÉRIE', '3ª SÉRIE'
])

# Agrupar por escola
df_noturno_escola = (
    df_noturno
    .groupby('INEP ESCOLA')
    .agg(
        DIREC=('DIREC', lambda x: x.mode().iloc[0]),
        MUNICÍPIO=('MUNICÍPIO', lambda x: x.mode().iloc[0]),
        ESCOLA=('ESCOLA', lambda x: x.mode().iloc[0]),
        **{
            'Professores EF-AF': ('Professores EF-AF', 'sum'),
            'Professores EM': ('Professores EM', 'sum')
        }
    )
    .reset_index()
)


# Base de dados do Saeb 2025 para ter as escolas prioritárias













































