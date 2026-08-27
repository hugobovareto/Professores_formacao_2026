'''
##### NOVA LÓGICA #####
Deixar 1 vaga para cada etapa e componente (se a escola ofertar). Ou seja:
- 1 vaga para LP para 9º ano EF.
- 1 vaga para MT para 9º ano EF.
- 1 vaga para LP para Ensino Médio.
- 1 vaga para MT para Ensino Médio.

1.791 vagas totais:
- 48 para as DIRECs;
- 30 para SEEC;

= 1.713 vagas para professores e coordenadores.
- 464 escolas, logo 464 vagas para coordenadores.

= 1.249 vagas para professores para dsitribuir entre os componentes e etapas.

'''


'''
##### LÓGICA ANTIGA #####
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
- 9º ano,
- 1ª série,
- 2ª série,
- 3ª série.

Considerar somente Ensino Regular (excluir EJA e EPT).

Para identificação de escolas com ensino noturno, considerar o código das turmas.


Regra para alocação das vagas:
1.791 vagas totais:
- 48 para as DIRECs;
- 30 para SEEC;

= 1.713 vagas para professores e coordenadores.

Dessas:
- 1 vaga para coordenador por escola que oferte 9º ano EF e Ensino Médio (464 vagas);

= 1.249 vagas para professores (diurnos e noturnos)
- 1 para professor de LP para cada etapa (Anos Finais EF e Ensino Médio) por escola;
- 1 para professor de LP para cada etapa (Anos Finais EF e Ensino Médio) por escola;
- o que sobra para noturno.

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

##### NOVA LÓGICA #####
# Relatório geral de matrículas - 2026
df_geral_26 = pd.read_excel(r"C:\Users\hugob\Downloads\20260826_Relatório Geral de Estudantes - Matrículas.xlsx", skiprows=2)

# DATAFRAMES RESERVAS
df_escolas = df_geral_26.copy(deep=True)

# Manter somente as Séries de interesse:
series = ['1ª SÉRIE',
          '2ª SÉRIE',
          '3ª SÉRIE',
          '9º ANO']

df_escolas = df_escolas[df_escolas['SÉRIE'].isin(series)]


# Manter o registro somente dos alunos Matriculados
df_escolas = df_escolas[df_escolas['SITUAÇÃO'].isin(['MATRICULADO'])]


# Contar a quantidade de escolas para saber as vagas de coordenadores
# Manter as colunas para a quantidade de escolas
df_escolas = df_escolas[['DIREC', 'CÓDIGO INEP ESCOLA', 'ESCOLA']]

# Excluir duplicatas de 'CÓDIGO INEP ESCOLA'
df_escolas = df_escolas.drop_duplicates(subset='CÓDIGO INEP ESCOLA', keep='first')

# Total de escolas: 464
len(df_escolas['CÓDIGO INEP ESCOLA'].unique())


# Criar base das escolas e indicar quais séries são ofertadas
# DATAFRAMES RESERVAS
df_professores = df_geral_26.copy(deep=True)

# Manter somente as Séries de interesse:
series = ['1ª SÉRIE',
          '2ª SÉRIE',
          '3ª SÉRIE',
          '9º ANO']

df_professores = df_professores[df_professores['SÉRIE'].isin(series)]


# Manter o registro somente dos alunos Matriculados
df_professores = df_professores[df_professores['SITUAÇÃO'].isin(['MATRICULADO'])]

# Manter somente as colunas de interesse
df_professores = df_professores[['DIREC', 'CÓDIGO INEP ESCOLA', 'ESCOLA', 'SÉRIE']]

# Agrupar por 'CÓDIGO INEP' e a 'SÉRIE' deve indicar todas as séries separadas por vírgula
df_professores_agrupado = (
    df_professores
    .groupby('CÓDIGO INEP ESCOLA', as_index=False)
    .agg({
        'DIREC': 'first',
        'ESCOLA': 'first',
        'SÉRIE': lambda x: ', '.join(x.astype(str).unique())
    })
)


# Criar a coluna 'OFERTA_9EF' para saber se a escola oferece 9º ano EF (se tem SÉRIE = '9º ANO')
df_professores_agrupado['OFERTA_9EF'] = np.where(
    df_professores_agrupado['SÉRIE'].str.contains('9º ANO', na=False),
    'Sim',
    'Não'
)


# Criar a coluna 'OFERTA_EM' para saber se a escola oferece Ensino Médio (se tem SÉRIE = '1ª SÉRIE' ou '2ª SÉRIE' ou '3ª SÉRIE')
df_professores_agrupado['OFERTA_EM'] = np.where(
    df_professores_agrupado['SÉRIE'].str.contains(
        '1ª SÉRIE|2ª SÉRIE|3ª SÉRIE',
        na=False
    ),
    'Sim',
    'Não'
)

# Criar coluna 'Vagas - Coordenadores' para indicar a vaga de 1 coordenador por escola
df_professores_agrupado['Vagas - Coordenadores'] = 1


# Criar as colunas de vagas para cada componente e cada etapa de ensino de acordo com os valores das colunas 'OFERTA_9EF' e 'OFERTA_EM'
# Vagas AF - LP; Vagas AF - MT; Vagas EM - LP; Vagas EM - MT
















##### Exportar para Excel
with pd.ExcelWriter('20260827_professores_formacao_2026.xlsx') as writer:
    df_professores_agrupado.to_excel(writer, sheet_name='Professores', index=False)
    df_escolas.to_excel(writer, sheet_name='Coordenadores_Escolas', index=False)



























##### Saber o número de escolas para saber a quantidade de vagas para coordenadores
##### LÓGICA ANTIGA #####

# Relatório geral de matrículas - 2026
df_geral_26 = pd.read_excel(r"C:\Users\hugob\Downloads\20260826_Relatório Geral de Estudantes - Matrículas.xlsx", skiprows=2)


# Manter somente as Séries de interesse:
series = ['1ª SÉRIE',
          '2ª SÉRIE',
          '3ª SÉRIE',
          '9º ANO']

df_geral_26 = df_geral_26[df_geral_26['SÉRIE'].isin(series)]

# NÃO TEVE EXCLUSÃO DE ETAPA, pois tem que incluir os cursos técnicos articulados e integrados ao Médio, pois esses estudantes tem ensino regular também
# Manter somente as etapas de ensino de interesse (excluir EPT e EJA)
# df_geral_26 = df_geral_26[df_geral_26['ETAPA DE ENSINO'].isin(['ENSINO MÉDIO POTIGUAR', 'ENSINO FUNDAMENTAL'])]

# Manter o registro somente dos alunos Matriculados
df_geral_26 = df_geral_26[df_geral_26['SITUAÇÃO'].isin(['MATRICULADO'])]


# Contar a quantidade de escolas para saber as vagas de coordenadores
# Manter as colunas para a quantidade de escolas
df_geral_26 = df_geral_26[['DIREC', 'CÓDIGO INEP ESCOLA', 'ESCOLA']]

# Excluir duplicatas de 'CÓDIGO INEP ESCOLA'
df_geral_26 = df_geral_26.drop_duplicates(subset='CÓDIGO INEP ESCOLA', keep='first')

# Total de escolas: 464
len(df_geral_26['CÓDIGO INEP ESCOLA'].unique())


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


# Manter somente as Séries de interesse (Anos Finais Ensino Fundamental e Ensino Médio)
series = ['1ª SÉRIE',
          '2ª SÉRIE',
          '3ª SÉRIE',
          '9º ANO']

df_freq_prof_26 = df_freq_prof_26[df_freq_prof_26['SÉRIE'].isin(series)]

# NÃO TEVE EXCLUSÃO DE ETAPA, pois tem que incluir os cursos técnicos articulados e integrados ao Médio, pois esses estudantes tem ensino regular também
# Manter somente as etapas de ensino de interesse (excluir EPT)
# df_freq_prof_26 = df_freq_prof_26[df_freq_prof_26['ETAPA DE ENSINO'].isin(['ENSINO MÉDIO POTIGUAR', 'ENSINO FUNDAMENTAL'])]


# Eliminar duplicatas de professores (MATRÍCULA) mantendo somente o maior número de 'AULAS PREVISTAS'
df_freq_prof_26 = df_freq_prof_26.sort_values(by=['MATRÍCULA', 'AULAS PREVISTAS'], ascending=False)
df_freq_prof_26 = df_freq_prof_26.drop_duplicates(subset=['MATRÍCULA'], keep='first')


# Adicionar coluna de Polo, de acordo com a DIREC
mapa_polos = {
    '01ª DIREC - NATAL': 'Polo 1',
    '05ª DIREC - CEARÁ MIRIM': 'Polo 1',

    '02ª DIREC - PARNAMIRIM': 'Polo 2',
    '06ª DIREC - MACAU': 'Polo 2',
    '04ª DIREC - SÃO PAULO DO POTENGI': 'Polo 2',
    '16ª DIREC - JOÃO CÂMARA': 'Polo 2',
    '03ª DIREC - NOVA CRUZ': 'Polo 2',

    '12ª DIREC - MOSSORÓ': 'Polo 3',
    '11ª DIREC - ASSU': 'Polo 3',
    '08ª DIREC - ANGICOS': 'Polo 3',

    '10ª DIREC - CAICÓ': 'Polo 4',
    '07ª DIREC - SANTA CRUZ': 'Polo 4',
    '09ª DIREC - CURRAIS NOVOS': 'Polo 4',

    '15ª DIREC - PAU DOS FERROS': 'Polo 5',
    '14ª DIREC - UMARIZAL': 'Polo 5',
    '13ª DIREC - APODI': 'Polo 5'
}

df_freq_prof_26['POLO'] = df_freq_prof_26['DIREC'].map(mapa_polos)

# São 1.276 vagas para os professores, incluindo noturno

len(df_freq_prof_26)
# Tem 1.647 professores na base, incluindo Noturno, então só 371 ficam de fora da Formação.


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
df_regular['Professores Anos Finais - LP'] = (df_regular['SÉRIE'].isin(['6º ANO', '7º ANO', '8º ANO', '9º ANO'])
& df_regular['COMPONENTE'].eq('Língua Portuguesa'))

df_regular['Professores Anos Finais - MT'] = (df_regular['SÉRIE'].isin(['6º ANO', '7º ANO', '8º ANO', '9º ANO'])
& df_regular['COMPONENTE'].eq('Matemática'))

df_regular['Professores Ensino Medio - LP'] = (df_regular['SÉRIE'].isin(['1ª SÉRIE', '2ª SÉRIE', '3ª SÉRIE'])
& df_regular['COMPONENTE'].eq('Língua Portuguesa'))

df_regular['Professores Ensino Medio - MT'] = (df_regular['SÉRIE'].isin(['1ª SÉRIE', '2ª SÉRIE', '3ª SÉRIE'])
& df_regular['COMPONENTE'].eq('Matemática'))


# Agrupar por escola
df_regular_escola = (
    df_regular
    .groupby('INEP ESCOLA')
    .agg(
        POLO=('POLO', lambda x: x.mode().iloc[0]),
        DIREC=('DIREC', lambda x: x.mode().iloc[0]),
        MUNICÍPIO=('MUNICÍPIO', lambda x: x.mode().iloc[0]),
        ESCOLA=('ESCOLA', lambda x: x.mode().iloc[0]),
        **{
            'Professores Anos Finais - LP': ('Professores Anos Finais - LP', 'sum'),
            'Professores Anos Finais - MT': ('Professores Anos Finais - MT', 'sum'),
            'Professores Ensino Medio - LP': ('Professores Ensino Medio - LP', 'sum'),
            'Professores Ensino Medio - MT': ('Professores Ensino Medio - MT', 'sum')
        }
    )
    .reset_index()
)


# NOTURNO
# Criar indicadores para cada linha
df_noturno['Professores Anos Finais - LP'] = (df_noturno['SÉRIE'].isin(['6º ANO', '7º ANO', '8º ANO', '9º ANO'])
& df_noturno['COMPONENTE'].eq('Língua Portuguesa'))

df_noturno['Professores Anos Finais - MT'] = (df_noturno['SÉRIE'].isin(['6º ANO', '7º ANO', '8º ANO', '9º ANO'])
& df_noturno['COMPONENTE'].eq('Matemática'))

df_noturno['Professores Ensino Medio - LP'] = (df_noturno['SÉRIE'].isin(['1ª SÉRIE', '2ª SÉRIE', '3ª SÉRIE'])
& df_noturno['COMPONENTE'].eq('Língua Portuguesa'))

df_noturno['Professores Ensino Medio - MT'] = (df_noturno['SÉRIE'].isin(['1ª SÉRIE', '2ª SÉRIE', '3ª SÉRIE'])
& df_noturno['COMPONENTE'].eq('Matemática'))


# Agrupar por escola
df_noturno_escola = (
    df_noturno
    .groupby('INEP ESCOLA')
    .agg(
        POLO=('POLO', lambda x: x.mode().iloc[0]),
        DIREC=('DIREC', lambda x: x.mode().iloc[0]),
        MUNICÍPIO=('MUNICÍPIO', lambda x: x.mode().iloc[0]),
        ESCOLA=('ESCOLA', lambda x: x.mode().iloc[0]),
        **{
            'Professores Anos Finais - LP': ('Professores Anos Finais - LP', 'sum'),
            'Professores Anos Finais - MT': ('Professores Anos Finais - MT', 'sum'),
            'Professores Ensino Medio - LP': ('Professores Ensino Medio - LP', 'sum'),
            'Professores Ensino Medio - MT': ('Professores Ensino Medio - MT', 'sum')
        }
    )
    .reset_index()
)


# Base de dados do Saeb 2025 para ter as escolas prioritárias
# Ler os dados de notas do Saeb 2025:
df_saeb_af = pd.read_excel(r"D:\Scripts_Python\FGV\Professores_Formacao_2026\Notas_Saeb_2025.xlsx", sheet_name= "Anos Finais")
df_saeb_em = pd.read_excel(r"D:\Scripts_Python\FGV\Professores_Formacao_2026\Notas_Saeb_2025.xlsx", sheet_name= "Ensino Medio")


# Adicionar coluna de nota do Saeb 2025 nos df_regular_escola e df_noturno_escola
# REGULAR
# Anos Finais
df_regular_escola['Saeb_2025_AF'] = (
    df_regular_escola['INEP ESCOLA']
    .map(
        df_saeb_af.set_index('ID_ESCOLA')['VL_NOTA_MEDIA_2025']
    )
)

# Ensino Médio
df_regular_escola['Saeb_2025_EM'] = (
    df_regular_escola['INEP ESCOLA']
    .map(
        df_saeb_em.set_index('ID_ESCOLA')['VL_NOTA_MEDIA_2025']
    )
)


# NOTURNO
# Anos Finais
df_noturno_escola['Saeb_2025_AF'] = (
    df_noturno_escola['INEP ESCOLA']
    .map(
        df_saeb_af.set_index('ID_ESCOLA')['VL_NOTA_MEDIA_2025']
    )
)

# Ensino Médio
df_noturno_escola['Saeb_2025_EM'] = (
    df_noturno_escola['INEP ESCOLA']
    .map(
        df_saeb_em.set_index('ID_ESCOLA')['VL_NOTA_MEDIA_2025']
    )
)


# Transformar as colunas de Saeb em número para possibilitar os cálculos
# REGULAR
df_regular_escola['Saeb_2025_AF'] = pd.to_numeric(
    df_regular_escola['Saeb_2025_AF'],
    errors='coerce'
)

df_regular_escola['Saeb_2025_EM'] = pd.to_numeric(
    df_noturno_escola['Saeb_2025_EM'],
    errors='coerce'
)

# NOTURNO
df_regular_escola['Saeb_2025_AF'] = pd.to_numeric(
    df_regular_escola['Saeb_2025_AF'],
    errors='coerce'
)

df_noturno_escola['Saeb_2025_EM'] = pd.to_numeric(
    df_noturno_escola['Saeb_2025_EM'],
    errors='coerce'
)

# 228 vagas para Noturno
# = 1.021 vagas para Diurno
# = 1.249 vagas totais para professores


############################################################################################ REGULAR ############################################################################################
# ==========================================
# 1. VAGAS GARANTIDAS
# ==========================================

df_regular_escola['Vagas AF - LP'] = (
    df_regular_escola['Professores Anos Finais - LP'] > 0
).astype(int)

df_regular_escola['Vagas AF - MT'] = (
    df_regular_escola['Professores Anos Finais - MT'] > 0
).astype(int)

df_regular_escola['Vagas EM - LP'] = (
    df_regular_escola['Professores Ensino Medio - LP'] > 0
).astype(int)

df_regular_escola['Vagas EM - MT'] = (
    df_regular_escola['Professores Ensino Medio - MT'] > 0
).astype(int)


# ==========================================
# 2. TOTAL INICIAL
# ==========================================

total_vagas_regular = (
    df_regular_escola['Vagas AF - LP'].sum()
    + df_regular_escola['Vagas AF - MT'].sum()
    + df_regular_escola['Vagas EM - LP'].sum()
    + df_regular_escola['Vagas EM - MT'].sum()
)

vagas_restantes_regular = 1021 - total_vagas_regular

print(f'Vagas garantidas: {total_vagas_regular}')
print(f'Vagas restantes: {vagas_restantes_regular}')


# ==========================================
# 3. DIVIDIR AS VAGAS RESTANTES ENTRE AF E EM
# ==========================================

# Se houver número ímpar de vagas,
# AF recebe a vaga adicional.

vagas_af = (vagas_restantes_regular + 1) // 2
vagas_em = vagas_restantes_regular // 2

print(f'Vagas adicionais AF: {vagas_af}')
print(f'Vagas adicionais EM: {vagas_em}')


# =============================================================================
# 4. ESCOLAS ELEGÍVEIS - ANOS FINAIS
# =============================================================================

# Escolas COM nota Saeb
# → serão atendidas primeiro
# → ordenadas da menor para a maior nota

df_prioridade_af = (
    df_regular_escola[
        df_regular_escola['Saeb_2025_AF'].notna() &
        (
            (df_regular_escola['Professores Anos Finais - LP'] > 0) |
            (df_regular_escola['Professores Anos Finais - MT'] > 0)
        )
    ]
    .sort_values('Saeb_2025_AF')
)


# Escolas SEM nota Saeb
# → só serão utilizadas depois que as escolas
#   com Saeb forem esgotadas

df_sem_saeb_af = (
    df_regular_escola[
        df_regular_escola['Saeb_2025_AF'].isna() &
        (
            (df_regular_escola['Professores Anos Finais - LP'] > 0) |
            (df_regular_escola['Professores Anos Finais - MT'] > 0)
        )
    ]
)


# =============================================================================
# 5. ESCOLAS ELEGÍVEIS - ENSINO MÉDIO
# =============================================================================

# Escolas COM nota Saeb
# → serão atendidas primeiro
# → ordenadas da menor para a maior nota

df_prioridade_em = (
    df_regular_escola[
        df_regular_escola['Saeb_2025_EM'].notna() &
        (
            (df_regular_escola['Professores Ensino Medio - LP'] > 0) |
            (df_regular_escola['Professores Ensino Medio - MT'] > 0)
        )
    ]
    .sort_values('Saeb_2025_EM')
)


# Escolas SEM nota Saeb
# → só serão utilizadas depois que as escolas
#   com Saeb forem esgotadas

df_sem_saeb_em = (
    df_regular_escola[
        df_regular_escola['Saeb_2025_EM'].isna() &
        (
            (df_regular_escola['Professores Ensino Medio - LP'] > 0) |
            (df_regular_escola['Professores Ensino Medio - MT'] > 0)
        )
    ]
)


# =============================================================================
# 6. FUNÇÃO PARA DISTRIBUIÇÃO DE VAGAS - ANOS FINAIS
# =============================================================================

def distribuir_vagas_af(df_escolas):

    global vagas_af

    while vagas_af > 0:

        vagas_distribuidas_rodada_af = 0

        # Percorre as escolas na ordem definida
        for idx in df_escolas.index:

            if vagas_af == 0:
                break

            # --------------------------------------
            # Verifica capacidade de LP
            # --------------------------------------

            pode_lp = (
                df_regular_escola.loc[idx, 'Vagas AF - LP']
                <
                df_regular_escola.loc[
                    idx,
                    'Professores Anos Finais - LP'
                ]
            )


            # --------------------------------------
            # Verifica capacidade de MT
            # --------------------------------------

            pode_mt = (
                df_regular_escola.loc[idx, 'Vagas AF - MT']
                <
                df_regular_escola.loc[
                    idx,
                    'Professores Anos Finais - MT'
                ]
            )


            # ======================================
            # LP E MT PODEM RECEBER
            # ======================================

            if pode_lp and pode_mt:

                # Se existem pelo menos 2 vagas,
                # uma vai para cada componente

                if vagas_af >= 2:

                    df_regular_escola.loc[
                        idx,
                        'Vagas AF - LP'
                    ] += 1

                    df_regular_escola.loc[
                        idx,
                        'Vagas AF - MT'
                    ] += 1

                    vagas_af -= 2

                    vagas_distribuidas_rodada_af += 2


                # ----------------------------------
                # Resta apenas 1 vaga
                # ----------------------------------

                else:

                    total_af_lp_atual = (
                        df_regular_escola['Vagas AF - LP'].sum()
                    )

                    total_af_mt_atual = (
                        df_regular_escola['Vagas AF - MT'].sum()
                    )


                    # A vaga vai para a área
                    # com menor quantidade acumulada

                    if total_af_lp_atual <= total_af_mt_atual:

                        df_regular_escola.loc[
                            idx,
                            'Vagas AF - LP'
                        ] += 1

                    else:

                        df_regular_escola.loc[
                            idx,
                            'Vagas AF - MT'
                        ] += 1


                    vagas_af -= 1

                    vagas_distribuidas_rodada_af += 1


            # ======================================
            # SOMENTE LP PODE RECEBER
            # ======================================

            elif pode_lp:

                df_regular_escola.loc[
                    idx,
                    'Vagas AF - LP'
                ] += 1

                vagas_af -= 1

                vagas_distribuidas_rodada_af += 1


            # ======================================
            # SOMENTE MT PODE RECEBER
            # ======================================

            elif pode_mt:

                df_regular_escola.loc[
                    idx,
                    'Vagas AF - MT'
                ] += 1

                vagas_af -= 1

                vagas_distribuidas_rodada_af += 1


        # ------------------------------------------
        # Nenhuma vaga foi distribuída nesta rodada
        # ------------------------------------------

        if vagas_distribuidas_rodada_af == 0:
            break


# =============================================================================
# 7. DISTRIBUIÇÃO AF
# =============================================================================

# PRIMEIRO:
# escolas com Saeb, da menor para a maior nota

distribuir_vagas_af(df_prioridade_af)


# DEPOIS:
# escolas sem Saeb

if vagas_af > 0:

    distribuir_vagas_af(df_sem_saeb_af)


# =============================================================================
# 8. FUNÇÃO PARA DISTRIBUIÇÃO DE VAGAS - ENSINO MÉDIO
# =============================================================================

def distribuir_vagas_em(df_escolas):

    global vagas_em

    while vagas_em > 0:

        vagas_distribuidas_rodada_em = 0

        # Percorre as escolas na ordem definida
        for idx in df_escolas.index:

            if vagas_em == 0:
                break

            # --------------------------------------
            # Verifica capacidade de LP
            # --------------------------------------

            pode_lp = (
                df_regular_escola.loc[idx, 'Vagas EM - LP']
                <
                df_regular_escola.loc[
                    idx,
                    'Professores Ensino Medio - LP'
                ]
            )


            # --------------------------------------
            # Verifica capacidade de MT
            # --------------------------------------

            pode_mt = (
                df_regular_escola.loc[idx, 'Vagas EM - MT']
                <
                df_regular_escola.loc[
                    idx,
                    'Professores Ensino Medio - MT'
                ]
            )


            # ======================================
            # LP E MT PODEM RECEBER
            # ======================================

            if pode_lp and pode_mt:

                # Se existem pelo menos 2 vagas,
                # uma vai para cada componente

                if vagas_em >= 2:

                    df_regular_escola.loc[
                        idx,
                        'Vagas EM - LP'
                    ] += 1

                    df_regular_escola.loc[
                        idx,
                        'Vagas EM - MT'
                    ] += 1

                    vagas_em -= 2

                    vagas_distribuidas_rodada_em += 2


                # ----------------------------------
                # Resta apenas 1 vaga
                # ----------------------------------

                else:

                    total_em_lp_atual = (
                        df_regular_escola['Vagas EM - LP'].sum()
                    )

                    total_em_mt_atual = (
                        df_regular_escola['Vagas EM - MT'].sum()
                    )


                    # A vaga vai para a área
                    # com menor quantidade acumulada

                    if total_em_lp_atual <= total_em_mt_atual:

                        df_regular_escola.loc[
                            idx,
                            'Vagas EM - LP'
                        ] += 1

                    else:

                        df_regular_escola.loc[
                            idx,
                            'Vagas EM - MT'
                        ] += 1


                    vagas_em -= 1

                    vagas_distribuidas_rodada_em += 1


            # ======================================
            # SOMENTE LP PODE RECEBER
            # ======================================

            elif pode_lp:

                df_regular_escola.loc[
                    idx,
                    'Vagas EM - LP'
                ] += 1

                vagas_em -= 1

                vagas_distribuidas_rodada_em += 1


            # ======================================
            # SOMENTE MT PODE RECEBER
            # ======================================

            elif pode_mt:

                df_regular_escola.loc[
                    idx,
                    'Vagas EM - MT'
                ] += 1

                vagas_em -= 1

                vagas_distribuidas_rodada_em += 1


        # ------------------------------------------
        # Nenhuma vaga foi distribuída nesta rodada
        # ------------------------------------------

        if vagas_distribuidas_rodada_em == 0:
            break


# =============================================================================
# 9. DISTRIBUIÇÃO EM
# =============================================================================

# PRIMEIRO:
# escolas com Saeb, da menor para a maior nota

distribuir_vagas_em(df_prioridade_em)


# DEPOIS:
# escolas sem Saeb

if vagas_em > 0:

    distribuir_vagas_em(df_sem_saeb_em)


# =============================================================================
# 10. CONFERÊNCIA FINAL
# =============================================================================

total_af_lp = df_regular_escola['Vagas AF - LP'].sum()
total_af_mt = df_regular_escola['Vagas AF - MT'].sum()

total_em_lp = df_regular_escola['Vagas EM - LP'].sum()
total_em_mt = df_regular_escola['Vagas EM - MT'].sum()


total_af = total_af_lp + total_af_mt
total_em = total_em_lp + total_em_mt

total_final_regular = total_af + total_em


print('--- ANOS FINAIS ---')
print(f'AF - LP: {total_af_lp}')
print(f'AF - MT: {total_af_mt}')
print(f'Total AF: {total_af}')

print()

print('--- ENSINO MÉDIO ---')
print(f'EM - LP: {total_em_lp}')
print(f'EM - MT: {total_em_mt}')
print(f'Total EM: {total_em}')

print()

print(f'Total final: {total_final_regular}')
print(f'Vagas restantes: {1021 - total_final_regular}')


# =============================================================================
# 11. CONFERÊNCIA DE RESPEITO AO LIMITE DE PROFESSORES
# =============================================================================

print()
print('--- CONFERÊNCIA DOS LIMITES ---')

print(
    'AF - LP respeita limite:',
    (
        df_regular_escola['Vagas AF - LP']
        <= df_regular_escola['Professores Anos Finais - LP']
    ).all()
)

print(
    'AF - MT respeita limite:',
    (
        df_regular_escola['Vagas AF - MT']
        <= df_regular_escola['Professores Anos Finais - MT']
    ).all()
)

print(
    'EM - LP respeita limite:',
    (
        df_regular_escola['Vagas EM - LP']
        <= df_regular_escola['Professores Ensino Medio - LP']
    ).all()
)

print(
    'EM - MT respeita limite:',
    (
        df_regular_escola['Vagas EM - MT']
        <= df_regular_escola['Professores Ensino Medio - MT']
    ).all()
)

############################################################################################ NOTURNO ############################################################################################
# ==========================================
# 1. VAGAS GARANTIDAS
# ==========================================

df_noturno_escola['Vagas EM - LP'] = (
    df_noturno_escola['Professores Ensino Medio - LP'] > 0
).astype(int)

df_noturno_escola['Vagas EM - MT'] = (
    df_noturno_escola['Professores Ensino Medio - MT'] > 0
).astype(int)


# ==========================================
# 2. TOTAL INICIAL
# ==========================================

total_vagas_noturno = (
    df_noturno_escola['Vagas EM - LP'].sum()
    + df_noturno_escola['Vagas EM - MT'].sum()
)

vagas_restantes_noturno = 228 - total_vagas_noturno

print(f'Vagas garantidas: {total_vagas_noturno}')
print(f'Vagas restantes: {vagas_restantes_noturno}')


# ==========================================
# 3. ESCOLAS COM SAEB
# ==========================================

# Escolas com nota Saeb
# → prioridade na distribuição
# → menor nota primeiro

df_prioridade_noturno = (
    df_noturno_escola[
        df_noturno_escola['Saeb_2025_EM'].notna() &
        (
            (df_noturno_escola['Professores Ensino Medio - LP'] > 0) |
            (df_noturno_escola['Professores Ensino Medio - MT'] > 0)
        )
    ]
    .sort_values('Saeb_2025_EM')
)


# ==========================================
# 4. ESCOLAS SEM SAEB
# ==========================================

# Essas escolas só receberão vagas depois
# que a capacidade das escolas com Saeb
# for esgotada.

df_sem_saeb_noturno = (
    df_noturno_escola[
        df_noturno_escola['Saeb_2025_EM'].isna() &
        (
            (df_noturno_escola['Professores Ensino Medio - LP'] > 0) |
            (df_noturno_escola['Professores Ensino Medio - MT'] > 0)
        )
    ]
)


# ==========================================
# 5. FUNÇÃO DE DISTRIBUIÇÃO
# ==========================================

def distribuir_vagas_noturno(df_escolas):

    global vagas_restantes_noturno

    while vagas_restantes_noturno > 0:

        vagas_distribuidas_rodada = 0

        # Percorre as escolas na ordem recebida
        for idx in df_escolas.index:

            if vagas_restantes_noturno == 0:
                break


            # --------------------------------------
            # Quantidade de professores
            # --------------------------------------

            professores_lp = df_noturno_escola.loc[
                idx,
                'Professores Ensino Medio - LP'
            ]

            professores_mt = df_noturno_escola.loc[
                idx,
                'Professores Ensino Medio - MT'
            ]


            # --------------------------------------
            # Quantidade de vagas já atribuídas
            # --------------------------------------

            vagas_lp = df_noturno_escola.loc[
                idx,
                'Vagas EM - LP'
            ]

            vagas_mt = df_noturno_escola.loc[
                idx,
                'Vagas EM - MT'
            ]


            # --------------------------------------
            # Verifica capacidade de cada área
            # --------------------------------------

            pode_lp = vagas_lp < professores_lp

            pode_mt = vagas_mt < professores_mt


            # ======================================
            # LP E MT PODEM RECEBER
            # ======================================

            if pode_lp and pode_mt:

                # ----------------------------------
                # Existem pelo menos 2 vagas
                # ----------------------------------

                if vagas_restantes_noturno >= 2:

                    df_noturno_escola.loc[
                        idx,
                        'Vagas EM - LP'
                    ] += 1

                    df_noturno_escola.loc[
                        idx,
                        'Vagas EM - MT'
                    ] += 1

                    vagas_restantes_noturno -= 2

                    vagas_distribuidas_rodada += 2


                # ----------------------------------
                # Existe somente 1 vaga
                # ----------------------------------

                else:

                    total_lp = (
                        df_noturno_escola['Vagas EM - LP'].sum()
                    )

                    total_mt = (
                        df_noturno_escola['Vagas EM - MT'].sum()
                    )


                    # A vaga vai para a área
                    # com menor quantidade acumulada

                    if total_lp <= total_mt:

                        df_noturno_escola.loc[
                            idx,
                            'Vagas EM - LP'
                        ] += 1

                    else:

                        df_noturno_escola.loc[
                            idx,
                            'Vagas EM - MT'
                        ] += 1


                    vagas_restantes_noturno -= 1

                    vagas_distribuidas_rodada += 1


            # ======================================
            # SOMENTE LP PODE RECEBER
            # ======================================

            elif pode_lp:

                df_noturno_escola.loc[
                    idx,
                    'Vagas EM - LP'
                ] += 1

                vagas_restantes_noturno -= 1

                vagas_distribuidas_rodada += 1


            # ======================================
            # SOMENTE MT PODE RECEBER
            # ======================================

            elif pode_mt:

                df_noturno_escola.loc[
                    idx,
                    'Vagas EM - MT'
                ] += 1

                vagas_restantes_noturno -= 1

                vagas_distribuidas_rodada += 1


        # ==========================================
        # SEGURANÇA CONTRA LOOP INFINITO
        # ==========================================

        if vagas_distribuidas_rodada == 0:
            break


# ==========================================
# 6. PRIMEIRA ETAPA:
#    ESCOLAS COM SAEB
# ==========================================

distribuir_vagas_noturno(
    df_prioridade_noturno
)


# ==========================================
# 7. SEGUNDA ETAPA:
#    ESCOLAS SEM SAEB
# ==========================================

if vagas_restantes_noturno > 0:

    distribuir_vagas_noturno(
        df_sem_saeb_noturno
    )


# ==========================================
# 8. CONFERÊNCIA FINAL
# ==========================================

total_lp_noturno = (
    df_noturno_escola['Vagas EM - LP'].sum()
)

total_mt_noturno = (
    df_noturno_escola['Vagas EM - MT'].sum()
)

total_final_noturno = (
    total_lp_noturno
    + total_mt_noturno
)


print('--- ENSINO MÉDIO NOTURNO ---')

print(f'Vagas EM - LP: {total_lp_noturno}')
print(f'Vagas EM - MT: {total_mt_noturno}')

print(f'Total final: {total_final_noturno}')

print(
    f'Vagas restantes: {228 - total_final_noturno}'
)


# ==========================================
# 9. CONFERÊNCIA DOS LIMITES
# ==========================================

print()
print('--- CONFERÊNCIA DOS LIMITES ---')

print(
    'EM - LP respeita limite:',
    (
        df_noturno_escola['Vagas EM - LP']
        <=
        df_noturno_escola['Professores Ensino Medio - LP']
    ).all()
)

print(
    'EM - MT respeita limite:',
    (
        df_noturno_escola['Vagas EM - MT']
        <=
        df_noturno_escola['Professores Ensino Medio - MT']
    ).all()
)

############################################### AGREGAÇÕES ###############################################
colunas_vagas = [
    'Vagas AF - LP',
    'Vagas AF - MT',
    'Vagas EM - LP',
    'Vagas EM - MT'
]

# Função para agregar por alguma coluna
def criar_agrupamento_vagas(df_regular, df_noturno, coluna_agrupamento):

    # -------------------------------
    # Regular
    # -------------------------------
    df_regular_temp = df_regular[
        [coluna_agrupamento] + colunas_vagas
    ].copy()


    # -------------------------------
    # Noturno
    # -------------------------------
    df_noturno_temp = df_noturno[
        [coluna_agrupamento, 'Vagas EM - LP', 'Vagas EM - MT']
    ].copy()

    # AF não existe no noturno → preencher com 0
    df_noturno_temp['Vagas AF - LP'] = 0
    df_noturno_temp['Vagas AF - MT'] = 0

    # Deixar as colunas na mesma ordem
    df_noturno_temp = df_noturno_temp[
        [coluna_agrupamento] + colunas_vagas
    ]


    # -------------------------------
    # Concatenar Regular + Noturno
    # -------------------------------
    df_temp = pd.concat(
        [df_regular_temp, df_noturno_temp],
        ignore_index=True
    )


    # -------------------------------
    # Agrupar e somar
    # -------------------------------
    df_resultado = (
        df_temp
        .groupby(coluna_agrupamento, as_index=False)[colunas_vagas]
        .sum()
    )


    return df_resultado


# Dataframe agrupado por DIREC
df_direc = criar_agrupamento_vagas(
    df_regular_escola,
    df_noturno_escola,
    'DIREC'
)


# Dataframe agrupado por POLO
df_polo = criar_agrupamento_vagas(
    df_regular_escola,
    df_noturno_escola,
    'POLO'
)


##### Exportar para Excel
with pd.ExcelWriter('professores_formacao_2026.xlsx') as writer:
    df_direc.to_excel(writer, sheet_name='Direc', index=False)
    df_polo.to_excel(writer, sheet_name='Polo', index=False)
    df_regular_escola.to_excel(writer, sheet_name='Escolas - Diurno', index=False)
    df_noturno_escola.to_excel(writer, sheet_name='Escolas - Noturno', index=False)
    df_freq_prof_26.to_excel(writer, sheet_name='Todos Professores', index=False)
    df_regular.to_excel(writer, sheet_name='Professores Diurno', index=False)
    df_noturno.to_excel(writer, sheet_name='Professores Noturno', index=False)
    df_geral_26.to_excel(writer, sheet_name='Coordenadores_Escolas', index=False)










