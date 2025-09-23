import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# Configuração do Google Sheets
def get_google_sheets():
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('google_credentials.json', scope)
    client = gspread.authorize(creds)
    return client

def atualizar_google_sheets(resultado, quarter, ano):
    try:
        client = get_google_sheets()
        sheet = client.open("D002 - DSE")
        
        # Buscar quarter nas abas
        for worksheet in [sheet.worksheet("Totais"), sheet.worksheet("Detalhes")]:
            existing_data = worksheet.get_all_records()
            
            # Verificar se o quarter já existe
            quarter_id = f"{quarter} {ano}"
            existing_quarters = [row['Quarter'] for row in existing_data]
            
            if worksheet.title == "Totais":
                new_row = [
                    quarter_id,
                    resultado["Total Local"],
                    resultado["Total Fora"],
                    resultado["Total Importado"]
                ]
            else:
                new_row = [
                    quarter_id,
                    resultado["Total Sucata"],
                    resultado["Total Beneficiamento"],
                    resultado["Total Geral"]
                ]
            
            if quarter_id in existing_quarters:
                # Encontrar linha para substituir
                row_index = existing_quarters.index(quarter_id) + 2  # +2 porque a linha 1 é cabeçalho
                worksheet.update(f"A{row_index}:D{row_index}", [new_row])
            else:
                worksheet.append_row(new_row)
        
        return True
    except Exception as e:
        st.error(f"Erro ao atualizar planilha: {str(e)}")
        return False

# Configurações gerais
st.set_page_config(page_title="Processador de Relatórios", page_icon=":bar_chart:", layout="wide")
pd.set_option("styler.render.max_elements", 1_000_000)

# Função para formatar números no padrão brasileiro
def formatar_numero_brasileiro(valor):
    """Formata números no padrão brasileiro."""
    if isinstance(valor, (int, float)):
        # Remove casas decimais e formata com pontos como separadores de milhares
        return f"{int(valor):,}".replace(',', '.')
    return valor

# Adicione este dicionário no início do código
MESES_POR_QUARTER = {
    1: {'quarter': 'Q1', 'meses': [1, 2, 3], 'nomes': ['Janeiro', 'Fevereiro', 'Março']},
    2: {'quarter': 'Q2', 'meses': [4, 5, 6], 'nomes': ['Abril', 'Maio', 'Junho']},
    3: {'quarter': 'Q3', 'meses': [7, 8, 9], 'nomes': ['Julho', 'Agosto', 'Setembro']},
    4: {'quarter': 'Q4', 'meses': [10, 11, 12], 'nomes': ['Outubro', 'Novembro', 'Dezembro']}
}

# Cabeçalho com logo e estilo
st.markdown(f"""
    <style>
        body {{
            background-color: #FFFFFF;
        }}
        .header {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 10px;
            border-bottom: 2px solid #01a9e0;
        }}
        .header img {{
            height: 60px;
        }}
        .header h1 {{
            color: #01a9e0;
            font-family: Arial, sans-serif;
            margin: 0;
        }}
        .kpi {{
            text-align: center;
            padding: 20px;
            margin: 10px;
            border-radius: 10px;
            background-color: #f9f9f9;
            border: 1px solid #ddd;
        }}
        .kpi h2 {{
            color: #01a9e0;
            font-size: 36px;
            margin-bottom: 10px;
        }}
        .kpi p {{
            font-size: 30px;
            color: #626366;
            font-weight: bold;
        }}
        .section-title {{
            margin-top: 40px;
            font-size: 28px;
            color: #626366;
            font-weight: bold;
        }}
        .separator {{
            border: 1px solid #ddd;
            margin: 20px 0;
        }}
    </style>
    <div class="header">
        <h1>Processador de Relatórios</h1>
        <img src="https://docol65anos.com.br/temp/logo-docol.png" alt="Logo da Empresa">
    </div>
""", unsafe_allow_html=True)

# Função para calcular desempenho trimestral
def calcular_desempenho_trimestral(df, selected_centers):
    df['Material'] = df['Material'].astype(str)
    df = df.dropna(subset=['Centro'])
    df['Centro'] = df['Centro'].astype(int)
    
    # Aplicar filtro de centros PRIMEIRO
    df_filtrado = df[df['Centro'].isin(selected_centers)].copy()
    
    # Identificar sucata 
    material_sucata_prefixos = ('10028330000', '10032677000', '100027000', '1000109000', '10001103000')
    df_sucata = df_filtrado[df_filtrado['Material'].str.startswith(material_sucata_prefixos)]
    soma_sucata = df_sucata['Quantidade'].astype(float).sum()  # Alterado para 'Quantidade'

    # Criar coluna 'Grupo'
    df_filtrado['Grupo'] = df_filtrado['Material'].str[:2]

    # Ajustar beneficiamento - Nova regra
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
    valor_local = df_validos[df_validos['UF'].isin(UF_validos)]['Quantidade'].astype(float).sum()  # Alterado para 'Quantidade'
    valor_fora = df_validos[~df_validos['UF'].isin(UF_validos)]['Quantidade'].astype(float).sum()  # Alterado para 'Quantidade'
    
    # Importação 
    valor_importado = df_filtrado[df_filtrado['UF'] == 'EX']['Quantidade'].astype(float).sum()  # Alterado para 'Quantidade'

    # Beneficiamento 
    valor_beneficiamento = df_filtrado[df_filtrado['Grupo'] == '80']['Quantidade'].astype(float).sum()  # Alterado para 'Quantidade'

    # Nacional maior que 500km
    valor_nacional_500km = df_validos[~df_validos['UF'].isin(UF_validos + ['EX'])]['Quantidade'].astype(float).sum()  # Alterado para 'Quantidade'

    # Soma total DEVE considerar apenas os centros selecionados
    total_geral = df_filtrado['Quantidade'].astype(float).sum()  # Alterado para 'Quantidade'

    # Porcentagens calculadas sobre o total filtrado
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

# Função para exibir KPIs
def exibir_kpi(label, valor, unidade=""):
    st.markdown(f"""
        <div class="kpi">
            <h2>{label}</h2>
            <p>{valor} {unidade}</p>  # Exibe o valor já formatado como string
        </div>
    """, unsafe_allow_html=True)

# Interface principal
st.title("Upload de Relatórios")
st.markdown("Faça o upload de três relatórios para análise trimestral.")

uploaded_files = [
    st.file_uploader(f"Envie o relatório do mês {i+1}", type=['xlsx'])
    for i in range(3)
]

if all(uploaded_files):
    
    # Widget para seleção de centros
    selected_centers = st.multiselect(
        "Selecione os centros para análise:",
        options=[1001, 1002, 1003],
        default=[1001, 1002]
    )

    dfs = [pd.read_excel(file) for file in uploaded_files]
    
    # Validação do Quarter
    meses = []
    anos = []
    arquivos_problema = []
    
    for i, df in enumerate(dfs):
        if 'Dt Lanct' in df.columns:
            df['Dt Lanct'] = pd.to_datetime(df['Dt Lanct'], errors='coerce')
            df = df.dropna(subset=['Dt Lanct'])
            
            if not df.empty:
                # Extrair mês e ano corretamente
                mes = df['Dt Lanct'].dt.day.mode()[0]
                ano = df['Dt Lanct'].dt.year.mode()[0]
                meses.append(mes)
                anos.append(ano)
            else:
                st.error(f"Arquivo {i+1} não contém datas válidas na coluna 'Dt Lanct'")
                st.stop()
        else:
            st.error(f"Arquivo {i+1} não possui coluna 'Dt Lanct'")
            st.stop()

     # Verificar consistência do ano
    if len(set(anos)) > 1:
        st.error("Arquivos de anos diferentes detectados!")
        st.stop()
    ano = anos[0] 

    # Determinar quarter
    quarter_encontrado = None
    for q in MESES_POR_QUARTER.values():
        if all(m in q['meses'] for m in meses):
            quarter_encontrado = q
            break

    if not quarter_encontrado:
        st.error("Arquivos não pertencem ao mesmo quarter!")
        quarters_detectados = set()
        for m in meses:
            for q in MESES_POR_QUARTER.values():
                if m in q['meses']:
                    quarters_detectados.add(q['quarter'])
        st.error(f"Detectado arquivos dos quarters: {', '.join(quarters_detectados)}")
        st.stop()

    # Verificar meses faltantes
    meses_faltantes = [m for m in quarter_encontrado['meses'] if m not in meses]
    if meses_faltantes:
        nomes_faltantes = [quarter_encontrado['nomes'][m - quarter_encontrado['meses'][0]] for m in meses_faltantes]
        st.error(f"Quarter {quarter_encontrado['quarter']} incompleto. Meses faltando: {', '.join(nomes_faltantes)}")
        st.stop()

    # Extrair e exibir o mês e o ano de cada relatório
    meses_anos = []
    for i, df in enumerate(dfs, start=1):
        if 'Dt Lanct' in df.columns:
            df['Dt Lanct'] = pd.to_datetime(df['Dt Lanct'], errors='coerce')
            if not df['Dt Lanct'].dropna().empty:
                mes_referencia = int(df['Dt Lanct'].dt.day.mode()[0])
                ano_referencia = int(df['Dt Lanct'].dt.year.mode()[0])
                meses_anos.append(f"{mes_referencia:02d}/{ano_referencia}")

    resultados_individuais = [calcular_desempenho_trimestral(df, selected_centers) for df in dfs]
    resultado_soma = calcular_desempenho_trimestral(pd.concat(dfs, ignore_index=True), selected_centers)

    # Exibir resultados individuais
    for i, (resultado, mes_ano) in enumerate(zip(resultados_individuais, meses_anos), start=1):
        st.markdown(f"### Resultado - {mes_ano}")

        col1, col2, col3 = st.columns(3)
        with col1:
            pass
        with col2:
            exibir_kpi("Total Geral", formatar_numero_brasileiro(resultado["Total Geral"]))
        with col3:
            pass

        col4, col5, col6 = st.columns(3)
        with col4:
            exibir_kpi("Total Local", formatar_numero_brasileiro(resultado["Total Local"]))
            exibir_kpi("% - Local", f"{resultado['% - Local']:.2%}")
        with col5:
            exibir_kpi("Total Fora", formatar_numero_brasileiro(resultado["Total Fora"]))
            exibir_kpi("% - Fora", f"{resultado['% - Fora']:.2%}")
        with col6:
            exibir_kpi("Total Importado", formatar_numero_brasileiro(resultado["Total Importado"]))
            exibir_kpi("% - Importação", f"{resultado['% - Importação']:.2%}")
        
        col7, col8, col9 = st.columns(3)
        with col7:
            exibir_kpi("Total Beneficiamento", formatar_numero_brasileiro(resultado["Total Beneficiamento"]))
            exibir_kpi("% - Beneficiamento", f"{resultado['% - Beneficiamento']:.2%}")
        with col8:
            exibir_kpi("Total Sucata", formatar_numero_brasileiro(resultado["Total Sucata"]))
            exibir_kpi("% - Sucata", f"{resultado['% - Sucata']:.2%}")
        with col9:
            exibir_kpi("Total Nacional Fora", formatar_numero_brasileiro(resultado["Total Nacional Fora"]))
            exibir_kpi("% - Nacional Fora", f"{resultado['% - Nacional Fora']:.2%}")
        
        st.markdown("---")

    st.markdown(f"### Total Consolidado {quarter_encontrado['quarter']} {ano}")

    # Botão para salvar no Google Sheets
    if st.button("📤 Salvar no Google Sheets"):
        ano = anos[0]
        quarter = quarter_encontrado['quarter']
        
        if atualizar_google_sheets(resultado_soma, quarter, ano):
            st.success("Dados salvos com sucesso no Google Sheets!")
        else:
            st.error("Falha ao salvar os dados")

    # Exibir resultados consolidados
    col1, col2, col3 = st.columns(3)
    with col1:
        pass
    with col2:
        exibir_kpi("Total Geral", formatar_numero_brasileiro(resultado_soma["Total Geral"]))
    with col3:
        pass

    col4, col5, col6 = st.columns(3)
    with col4:
        exibir_kpi("Total Local", formatar_numero_brasileiro(resultado_soma["Total Local"]))
        exibir_kpi("% - Local", f"{resultado_soma['% - Local']:.2%}")
    with col5:
        exibir_kpi("Total Fora", formatar_numero_brasileiro(resultado_soma["Total Fora"]))
        exibir_kpi("% - Fora", f"{resultado_soma['% - Fora']:.2%}")
    with col6:
        exibir_kpi("Total Importado", formatar_numero_brasileiro(resultado_soma["Total Importado"]))
        exibir_kpi("% - Importação", f"{resultado_soma['% - Importação']:.2%}")
    
    col7, col8, col9 = st.columns(3)
    with col7:
        exibir_kpi("Total Beneficiamento", formatar_numero_brasileiro(resultado_soma["Total Beneficiamento"]))
        exibir_kpi("% - Beneficiamento", f"{resultado_soma['% - Beneficiamento']:.2%}")
    with col8:
        exibir_kpi("Total Sucata", formatar_numero_brasileiro(resultado_soma["Total Sucata"]))
        exibir_kpi("% - Sucata", f"{resultado_soma['% - Sucata']:.2%}")
    with col9:
        exibir_kpi("Total Nacional Fora", formatar_numero_brasileiro(resultado_soma["Total Nacional Fora"]))
        exibir_kpi("% - Nacional Fora", f"{resultado_soma['% - Nacional Fora']:.2%}")
    
    st.markdown("---")