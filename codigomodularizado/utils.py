"""
Módulo de funções utilitárias para o processador de relatórios.
Inclui cálculos, formatação e salvamento de arquivos.
"""
import streamlit as st
import pandas as pd
import re
import io  # Necessário para manipulação de arquivos em memória

# --- Constantes ---
MESES_POR_QUARTER = {
    1: {'quarter': 'Q1', 'meses': [1, 2, 3], 'nomes': ['Janeiro', 'Fevereiro', 'Março']},
    2: {'quarter': 'Q2', 'meses': [4, 5, 6], 'nomes': ['Abril', 'Maio', 'Junho']},
    3: {'quarter': 'Q3', 'meses': [7, 8, 9], 'nomes': ['Julho', 'Agosto', 'Setembro']},
    4: {'quarter': 'Q4', 'meses': [10, 11, 12], 'nomes': ['Outubro', 'Novembro', 'Dezembro']}
}

# --- NOVA FUNÇÃO: Salvar Relatório em Excel ---
def salvar_relatorio_excel(resultados_individuais, resultado_soma, meses_anos, quarter_info, ano_analise):
    """
    Cria um arquivo Excel em memória com 4 abas: uma para cada mês e uma para o consolidado.
    Retorna o arquivo como um objeto binário para download.
    """
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 1. Cria uma aba para cada resultado mensal
        for resultado, mes_ano in zip(resultados_individuais, meses_anos):
            df_mes = pd.DataFrame(list(resultado.items()), columns=['Indicador', 'Valor'])
            sheet_name = mes_ano.replace('/', '-')
            df_mes.to_excel(writer, sheet_name=sheet_name, index=False)

        # 2. Cria a aba para o resultado consolidado
        df_consolidado = pd.DataFrame(list(resultado_soma.items()), columns=['Indicador', 'Valor'])
        sheet_name_consolidado = f"Consolidado {quarter_info}-{ano_analise}"
        df_consolidado.to_excel(writer, sheet_name=sheet_name_consolidado, index=False)

    data = output.getvalue()
    return data

# --- Função de Conversão ---
def carregar_e_preparar_conversao(caminho_do_arquivo):
    try:
        df_conversao = pd.read_excel(caminho_do_arquivo, engine='openpyxl')
        df_conversao = df_conversao.dropna(subset=['Material'])
        df_conversao['Codigo_Material'] = df_conversao['Material'].str.extract(r'\((\d{8,})\)').astype(str)
        df_fatores = df_conversao[['Codigo_Material', 'Conversão na unidade de medida básica', 'Unidade de medida do pedido']].copy()
        df_fatores = df_fatores.dropna(subset=['Codigo_Material'])
        df_fatores.rename(columns={'Conversão na unidade de medida básica': 'Fator_Conversao'}, inplace=True)
        df_fatores = df_fatores.drop_duplicates(subset=['Codigo_Material'])
        mapa_moda = df_conversao.groupby('Unidade de medida do pedido')['Conversão na unidade de medida básica'].apply(lambda x: x.mode().iloc[0] if not x.mode().empty else 1).to_dict()
        return df_fatores, mapa_moda
    except FileNotFoundError:
        st.error(f"Arquivo de conversão não encontrado em: {caminho_do_arquivo}")
        return None, None
    except Exception as e:
        st.error(f"Erro ao carregar ou processar o arquivo de conversão: {e}")
        return None, None

# --- Funções de Formatação e Exibição ---
def formatar_numero_financeiro(valor):
    if isinstance(valor, (int, float)):
        return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    return valor

def formatar_numero_quantidade(valor):
    if isinstance(valor, (int, float)):
        return f"{int(valor):,}".replace(',', '.')
    return valor

def exibir_kpi(label, valor, unidade=""):
    st.markdown(f'<div class="kpi"><h2>{label}</h2><p>{valor} {unidade}</p></div>', unsafe_allow_html=True)

# --- Lógica de Cálculo Principal ---
def calcular_desempenho(df, selected_centers, coluna_calculo, df_conversao=None, mapa_moda_conversao=None):
    df['Material'] = df['Material'].astype(str)
    df = df.dropna(subset=['Centro'])
    df['Centro'] = df['Centro'].astype(int)
    df_filtrado = df[df['Centro'].isin(selected_centers)].copy()
    if coluna_calculo == 'Quantidade' and df_conversao is not None and mapa_moda_conversao is not None:
        df_filtrado[coluna_calculo] = pd.to_numeric(df_filtrado[coluna_calculo], errors='coerce').fillna(0)
        df_merged = pd.merge(df_filtrado, df_conversao, left_on='Material', right_on='Codigo_Material', how='left')
        nao_encontrados_mask = df_merged['Fator_Conversao'].isnull()
        coluna_unidade_medida_principal = 'Unidade de medida do pedido'
        if coluna_unidade_medida_principal in df_merged.columns:
            df_merged.loc[nao_encontrados_mask, 'Fator_Conversao'] = df_merged.loc[nao_encontrados_mask, coluna_unidade_medida_principal].map(mapa_moda_conversao)
        df_merged['Fator_Conversao'].fillna(1, inplace=True)
        df_merged[coluna_calculo] = df_merged[coluna_calculo] * df_merged['Fator_Conversao']
        df_filtrado = df_merged
    material_sucata_prefixos = ('10028330000', '10032677000', '100027000', '1000109000', '10001103000')
    df_sucata = df_filtrado[df_filtrado['Material'].str.lstrip('0').str.startswith(material_sucata_prefixos)]
    soma_sucata = df_sucata[coluna_calculo].astype(float).sum()
    df_filtrado['Grupo'] = df_filtrado['Material'].str[:2]
    condicao_beneficiamento = (df_filtrado['Cliente/Fornec'] == 1048374) & (df_filtrado['Código do IVA'] == 'I2')
    df_filtrado.loc[condicao_beneficiamento, 'Grupo'] = '80'
    grupos_validos = ['10', '11', '30']
    df_validos = df_filtrado[df_filtrado['Grupo'].isin(grupos_validos)]
    UF_validos = ['SC', 'PR']
    valor_local = df_validos[df_validos['UF'].isin(UF_validos)][coluna_calculo].astype(float).sum()
    valor_fora = df_validos[~df_validos['UF'].isin(UF_validos)][coluna_calculo].astype(float).sum()
    valor_importado = df_filtrado[df_filtrado['UF'] == 'EX'][coluna_calculo].astype(float).sum()
    valor_beneficiamento = df_filtrado[df_filtrado['Grupo'] == '80'][coluna_calculo].astype(float).sum()
    valor_nacional_500km = df_validos[~df_validos['UF'].isin(UF_validos + ['EX'])][coluna_calculo].astype(float).sum()
    total_geral = df_filtrado[coluna_calculo].astype(float).sum()
    porcentagens = {"% - Sucata": soma_sucata / total_geral if total_geral else 0,"% - Beneficiamento": valor_beneficiamento / total_geral if total_geral else 0,"% - Local": valor_local / total_geral if total_geral else 0,"% - Fora": valor_fora / total_geral if total_geral else 0,"% - Importação": valor_importado / total_geral if total_geral else 0,"% - Nacional Fora": valor_nacional_500km / total_geral if total_geral else 0,}
    return {"Total Local": valor_local, "Total Fora": valor_fora, "Total Importado": valor_importado, "Total Beneficiamento": valor_beneficiamento, "Total Sucata": soma_sucata, "Total Nacional Fora": valor_nacional_500km, "Total Geral": total_geral, **porcentagens}

