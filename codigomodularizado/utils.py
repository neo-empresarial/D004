# Módulo de funções utilitárias para o processador de relatórios.
import streamlit as st
import pandas as pd
import re
import io
import numpy as np

# --- Constantes ---
MESES_POR_QUARTER = {
    1: {'quarter': 'Q1', 'meses': [1, 2, 3], 'nomes': ['Janeiro', 'Fevereiro', 'Março']},
    2: {'quarter': 'Q2', 'meses': [4, 5, 6], 'nomes': ['Abril', 'Maio', 'Junho']},
    3: {'quarter': 'Q3', 'meses': [7, 8, 9], 'nomes': ['Julho', 'Agosto', 'Setembro']},
    4: {'quarter': 'Q4', 'meses': [10, 11, 12], 'nomes': ['Outubro', 'Novembro', 'Dezembro']}
}

# --- FUNÇÃO DE SALVAR RELATÓRIO (ATUALIZADA) ---
def salvar_relatorio_excel(resultados_individuais, resultado_soma, meses_anos, quarter_info, ano_analise, nome_analise_arquivo):
    """
    Salva um relatório em Excel formatado corretamente.
    - Altera o nome da coluna de valor para "Quantidade" se necessário.
    - Formata as células de porcentagem.
    """
    output = io.BytesIO()
    
    # Define o nome da coluna de valor com base no tipo de análise
    nome_coluna_valor = "Quantidade" if nome_analise_arquivo == "Quantidade" else "Valor"

    kpis_desejados = [
        "Total Local", "Total Fora", "Total Importado", "Total Beneficiamento",
        "Total Sucata", "Total Nacional Fora", "Total Geral", "% - Sucata",
        "% - Beneficiamento", "% - Local", "% - Fora", "% - Importação",
        "% - Nacional Fora"
    ]

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        # Cria um formato para porcentagem
        percent_format = workbook.add_format({'num_format': '0.00%'})

        # Salva a aba de cada mês
        for resultado, mes_ano in zip(resultados_individuais, meses_anos):
            resultado_filtrado = {key: resultado[key] for key in kpis_desejados if key in resultado}
            # Usa o nome da coluna correto
            df_mes = pd.DataFrame(list(resultado_filtrado.items()), columns=['Indicador', nome_coluna_valor])
            
            sheet_name = mes_ano.replace('/', '-')
            df_mes.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Aplica a formatação de porcentagem na aba recém-criada
            worksheet = writer.sheets[sheet_name]
            # Itera pelas linhas do DataFrame para encontrar os indicadores de porcentagem
            for row_num, indicador in enumerate(df_mes['Indicador']):
                if indicador.startswith('%'):
                    # A linha no Excel é row_num + 1 (cabeçalho)
                    # A coluna de valor é a segunda (índice 1)
                    worksheet.write(row_num + 1, 1, df_mes.iloc[row_num, 1], percent_format)

        # Salva a aba do consolidado
        resultado_soma_filtrado = {key: resultado_soma[key] for key in kpis_desejados if key in resultado_soma}
        # Usa o nome da coluna correto
        df_consolidado = pd.DataFrame(list(resultado_soma_filtrado.items()), columns=['Indicador', nome_coluna_valor])
        
        sheet_name_consolidado = f"Consolidado {quarter_info}-{ano_analise}"
        df_consolidado.to_excel(writer, sheet_name=sheet_name_consolidado, index=False)

        # Aplica a formatação de porcentagem na aba do consolidado
        worksheet_consolidado = writer.sheets[sheet_name_consolidado]
        for row_num, indicador in enumerate(df_consolidado['Indicador']):
            if indicador.startswith('%'):
                worksheet_consolidado.write(row_num + 1, 1, df_consolidado.iloc[row_num, 1], percent_format)


    return output.getvalue()


# --- Função de Conversão ---
def carregar_e_preparar_conversao(arquivo_carregado):
    try:
        df = pd.read_csv(arquivo_carregado) if arquivo_carregado.name.endswith('.csv') else pd.read_excel(arquivo_carregado)
        colunas_necessarias = ['Material', 'Conversão na unidade de medida básica']
        if not all(col in df.columns for col in colunas_necessarias):
            raise ValueError("O arquivo de conversão não contém as colunas 'Material' e 'Conversão na unidade de medida básica'.")
        df = df.dropna(subset=colunas_necessarias)
        df['Codigo_Material'] = df['Material'].str.extract(r'\((\d+)\)').astype(str)
        df['Codigo_Material'] = df['Codigo_Material'].str.lstrip('0')
        df = df.dropna(subset=['Codigo_Material'])
        mapa_conversao = pd.Series(
            df['Conversão na unidade de medida básica'].values,
            index=df['Codigo_Material']
        ).to_dict()
        return mapa_conversao
    except Exception as e:
        raise Exception(f"Falha ao processar o arquivo de conversão: {e}")


# --- Funções de Formatação e Exibição ---
def formatar_numero_financeiro(valor):
    if isinstance(valor, (int, float)): return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    return valor

def formatar_numero_quantidade(valor):
    if isinstance(valor, (int, float)): return f"{int(valor):,}".replace(',', '.')
    return valor

def exibir_kpi(label, valor, unidade=""):
    st.markdown(f'<div class="kpi"><h2>{label}</h2><p>{valor} {unidade}</p></div>', unsafe_allow_html=True)

# --- Lógica de Cálculo Principal ---
def calcular_desempenho(df, selected_centers, coluna_calculo, mapa_conversao=None, debug_contexto=""):
    # (O restante do arquivo utils.py continua exatamente o mesmo)
    df['Material'] = df['Material'].astype(str).str.strip()
    df = df.dropna(subset=['Centro'])
    df['Centro'] = df['Centro'].astype(int)
    df_filtrado = df[df['Centro'].isin(selected_centers)].copy()
    
    df_filtrado['Material'] = df_filtrado['Material'].str.lstrip('0')
    
    if coluna_calculo == 'Quantidade' and mapa_conversao is not None:
        materiais_faltantes = set()
        if 'Unidade de medida' not in df_filtrado.columns:
            return {
                'status': 'erro_conversao',
                'materiais_faltantes': [('N/A', "A coluna 'Unidade de medida' não foi encontrada nos relatórios.", 'N/A')]
            }
        df_filtrado['Unidade de medida'] = df_filtrado['Unidade de medida'].astype(str).str.strip().str.lower()
        mapa_simples = {
            't': 1000, 'tonelada': 1000,
            'milhar': 1000, 'milha': 1000,
            'cento': 100,
            'ml': 1000, 'pt': 1000,
        }
        unidades_complexas = ['cx', 'caixa', 'cj', 'conjunto', 'pac', 'pct', 'pacote', 'rl', 'rolo', 'sac', 'saco']
        def converter_quantidade(row):
            unidade = row['Unidade de medida']
            material = row['Material']
            quantidade = row['Quantidade']
            if unidade in mapa_simples:
                return quantidade * mapa_simples[unidade]
            elif unidade in unidades_complexas:
                fator = mapa_conversao.get(material)
                if fator is not None and pd.notna(fator):
                    return quantidade * fator
                else:
                    materiais_faltantes.add((material, row.get('Descrição do item', 'N/A'), row['Unidade de medida'].upper()))
                    return quantidade
            return quantidade
        df_filtrado['Quantidade'] = df_filtrado.apply(converter_quantidade, axis=1)
        if materiais_faltantes:
            return {
                'status': 'erro_conversao',
                'materiais_faltantes': sorted(list(materiais_faltantes))
            }

    df_filtrado['Categoria'] = 'Indefinido'
    material_sucata_prefixos = ('10028330000', '10032677000', '10002709000', '10001099000', '10001103000')
    mask_sucata = df_filtrado['Material'].str.startswith(material_sucata_prefixos)
    cliente_fornec_numeric = pd.to_numeric(df_filtrado['Cliente/Fornec'], errors='coerce')
    mask_benef = (cliente_fornec_numeric == 1048374) & (df_filtrado['Código do IVA'] == 'I2')
    mask_importado = df_filtrado['UF'] == 'EX'
    mask_local = df_filtrado['UF'].isin(['SC', 'PR'])
    mask_uf_nula = df_filtrado['UF'].isnull() | (df_filtrado['UF'] == '')

    df_filtrado.loc[mask_sucata, 'Categoria'] = 'Sucata'
    df_filtrado.loc[mask_benef & (df_filtrado['Categoria'] == 'Indefinido'), 'Categoria'] = 'Beneficiamento'
    df_filtrado.loc[mask_importado & (df_filtrado['Categoria'] == 'Indefinido'), 'Categoria'] = 'Importado'
    df_filtrado.loc[mask_uf_nula & (df_filtrado['Categoria'] == 'Indefinido'), 'Categoria'] = 'UF Nula'
    df_filtrado.loc[mask_local & (df_filtrado['Categoria'] == 'Indefinido'), 'Categoria'] = 'Local'
    df_filtrado.loc[df_filtrado['Categoria'] == 'Indefinido', 'Categoria'] = 'Fora'

    soma_sucata = df_filtrado.loc[df_filtrado['Categoria'] == 'Sucata', coluna_calculo].sum()
    valor_beneficiamento = df_filtrado.loc[df_filtrado['Categoria'] == 'Beneficiamento', coluna_calculo].sum()
    valor_importado = df_filtrado.loc[df_filtrado['Categoria'] == 'Importado', coluna_calculo].sum()
    valor_local = df_filtrado.loc[df_filtrado['Categoria'] == 'Local', coluna_calculo].sum()
    valor_fora = df_filtrado.loc[df_filtrado['Categoria'] == 'Fora', coluna_calculo].sum()
    valor_nacional_500km = valor_fora
    total_geral = df_filtrado[coluna_calculo].sum()
    porcentagens = {
        "% - Sucata": soma_sucata / total_geral if total_geral else 0,
        "% - Beneficiamento": valor_beneficiamento / total_geral if total_geral else 0,
        "% - Local": valor_local / total_geral if total_geral else 0,
        "% - Fora": valor_fora / total_geral if total_geral else 0,
        "% - Importação": valor_importado / total_geral if total_geral else 0,
        "% - Nacional Fora": valor_nacional_500km / total_geral if total_geral else 0
    }
    
    return {
        "status": "sucesso",
        "Total Local": valor_local, "Total Fora": valor_fora, "Total Importado": valor_importado, 
        "Total Beneficiamento": valor_beneficiamento, "Total Sucata": soma_sucata, 
        "Total Nacional Fora": valor_nacional_500km, "Total Geral": total_geral, 
        **porcentagens,
    }