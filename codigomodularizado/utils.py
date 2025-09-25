"""
Módulo de funções utilitárias para o processador de relatórios.
Inclui cálculos, formatação e comunicação com APIs.
"""
import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- Constantes ---
MESES_POR_QUARTER = {
    1: {'quarter': 'Q1', 'meses': [1, 2, 3], 'nomes': ['Janeiro', 'Fevereiro', 'Março']},
    2: {'quarter': 'Q2', 'meses': [4, 5, 6], 'nomes': ['Abril', 'Maio', 'Junho']},
    3: {'quarter': 'Q3', 'meses': [7, 8, 9], 'nomes': ['Julho', 'Agosto', 'Setembro']},
    4: {'quarter': 'Q4', 'meses': [10, 11, 12], 'nomes': ['Outubro', 'Novembro', 'Dezembro']}
}

# --- Funções de Formatação e Exibição ---

def formatar_numero_financeiro(valor):
    """Formata números no padrão monetário brasileiro (R$)."""
    if isinstance(valor, (int, float)):
        return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    return valor

def formatar_numero_quantidade(valor):
    """Formata números para quantidade, sem casas decimais."""
    if isinstance(valor, (int, float)):
        return f"{int(valor):,}".replace(',', '.')
    return valor

def exibir_kpi(label, valor, unidade=""):
    """Exibe um card de KPI formatado."""
    st.markdown(f"""
        <div class="kpi">
            <h2>{label}</h2>
            <p>{valor} {unidade}</p>
        </div>
    """, unsafe_allow_html=True)

# --- Lógica de Cálculo Principal ---

def calcular_desempenho(df, selected_centers, coluna_calculo):
    """
    Calcula o desempenho trimestral com base em uma coluna específica.
    
    Args:
        df (pd.DataFrame): O DataFrame com os dados.
        selected_centers (list): Lista de centros selecionados para o filtro.
        coluna_calculo (str): O nome da coluna para os cálculos ('Valor líquido' ou 'Quantidade').
    """
    df['Material'] = df['Material'].astype(str)
    df = df.dropna(subset=['Centro'])
    df['Centro'] = df['Centro'].astype(int)
    
    df_filtrado = df[df['Centro'].isin(selected_centers)].copy()
    
    # Identificar sucata (lógica dos zeros à esquerda já aplicada)
    material_sucata_prefixos = ('10028330000', '10032677000', '100027000', '1000109000', '10001103000')
    df_sucata = df_filtrado[df_filtrado['Material'].str.lstrip('0').str.startswith(material_sucata_prefixos)]
    soma_sucata = df_sucata[coluna_calculo].astype(float).sum()

    # Criar coluna 'Grupo'
    df_filtrado['Grupo'] = df_filtrado['Material'].str[:2]

    # Ajustar beneficiamento
    condicao_beneficiamento = (
        (df_filtrado['Cliente/Fornec'] == 1048374) & 
        (df_filtrado['Código do IVA'] == 'I2')
    )
    df_filtrado.loc[condicao_beneficiamento, 'Grupo'] = '80'

    # Filtrar grupos válidos
    grupos_validos = ['10', '11', '30']
    df_validos = df_filtrado[df_filtrado['Grupo'].isin(grupos_validos)]

    # Localização
    UF_validos = ['SC', 'PR']
    valor_local = df_validos[df_validos['UF'].isin(UF_validos)][coluna_calculo].astype(float).sum()
    valor_fora = df_validos[~df_validos['UF'].isin(UF_validos)][coluna_calculo].astype(float).sum()
    
    # Importação
    valor_importado = df_filtrado[df_filtrado['UF'] == 'EX'][coluna_calculo].astype(float).sum()

    # Beneficiamento
    valor_beneficiamento = df_filtrado[df_filtrado['Grupo'] == '80'][coluna_calculo].astype(float).sum()

    # Nacional maior que 500km
    valor_nacional_500km = df_validos[~df_validos['UF'].isin(UF_validos + ['EX'])][coluna_calculo].astype(float).sum()

    # Soma total
    total_geral = df_filtrado[coluna_calculo].astype(float).sum()

    # Porcentagens
    porcentagens = {
        "% - Sucata": soma_sucata / total_geral if total_geral else 0,
        "% - Beneficiamento": valor_beneficiamento / total_geral if total_geral else 0,
        "% - Local": valor_local / total_geral if total_geral else 0,
        "% - Fora": valor_fora / total_geral if total_geral else 0,
        "% - Importação": valor_importado / total_geral if total_geral else 0,
        "% - Nacional Fora": valor_nacional_500km / total_geral if total_geral else 0,
    }

    return {
        "Total Local": valor_local,
        "Total Fora": valor_fora,
        "Total Importado": valor_importado,
        "Total Beneficiamento": valor_beneficiamento,
        "Total Sucata": soma_sucata,
        "Total Nacional Fora": valor_nacional_500km,
        "Total Geral": total_geral,
        **porcentagens,
    }

# --- Funções de Integração (Google Sheets) ---

def get_google_sheets_client():
    """Retorna um cliente autorizado para interagir com o Google Sheets."""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('google_credentials.json', scope)
        client = gspread.authorize(creds)
        return client
    except FileNotFoundError:
        st.error("Arquivo 'google_credentials.json' não encontrado. Verifique se ele está no diretório correto.")
        return None
    except Exception as e:
        st.error(f"Erro ao autenticar com o Google Sheets: {e}")
        return None

def atualizar_google_sheets(resultado, quarter, ano):
    """Atualiza a planilha do Google Sheets com os resultados consolidados."""
    client = get_google_sheets_client()
    if not client:
        return False
        
    try:
        sheet = client.open("D002 - DSE")
        
        for worksheet in [sheet.worksheet("Totais"), sheet.worksheet("Detalhes")]:
            existing_data = worksheet.get_all_records()
            quarter_id = f"{quarter} {ano}"
            existing_quarters = [row.get('Quarter', '') for row in existing_data]
            
            if worksheet.title == "Totais":
                new_row = [quarter_id, resultado["Total Local"], resultado["Total Fora"], resultado["Total Importado"]]
            else:
                new_row = [quarter_id, resultado["Total Sucata"], resultado["Total Beneficiamento"], resultado["Total Geral"]]
            
            if quarter_id in existing_quarters:
                row_index = existing_quarters.index(quarter_id) + 2
                worksheet.update(f"A{row_index}:D{row_index}", [new_row])
            else:
                worksheet.append_row(new_row)
        
        return True
    except gspread.exceptions.SpreadsheetNotFound:
        st.error("Planilha 'D002 - DSE' não encontrada. Verifique o nome e as permissões.")
        return False
    except gspread.exceptions.WorksheetNotFound as e:
        st.error(f"Aba não encontrada na planilha: {e}. Verifique se as abas 'Totais' e 'Detalhes' existem.")
        return False
    except Exception as e:
        st.error(f"Erro desconhecido ao atualizar a planilha: {e}")
        return False
