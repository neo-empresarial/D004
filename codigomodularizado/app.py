"""
Aplicação principal Streamlit para processamento de relatórios trimestrais.
Permite a análise por valor financeiro ou por quantidade.
"""
import streamlit as st
import pandas as pd
import time
from style import get_css_block, get_header_html
from utils import (
    calcular_desempenho,
    formatar_numero_financeiro,
    formatar_numero_quantidade,
    exibir_kpi,
    atualizar_google_sheets,
    MESES_POR_QUARTER
)

# --- Configuração da Página ---
st.set_page_config(page_title="Processador de Relatórios Docol", page_icon="📊", layout="wide")
st.markdown(get_css_block(), unsafe_allow_html=True)
st.markdown(get_header_html(), unsafe_allow_html=True)
pd.set_option("styler.render.max_elements", 1_000_000)

# --- Interface do Usuário (Seleção e Upload) ---
st.markdown("### Selecione o tipo de análise:")
tipo_analise = st.selectbox(
    "Selecione o tipo de análise:",
    ("Análise Financeira (R$)", "Análise por Quantidade"),
    label_visibility="collapsed" # Oculta o label pois já temos um título
)

st.markdown("### Faça o upload dos três relatórios para a análise trimestral:")
uploaded_files = [
    st.file_uploader(f"Relatório do mês {i+1}", type=['xlsx', 'csv'], key=f"file{i}")
    for i in range(3)
]

# A lógica da aplicação só continua se todos os arquivos forem enviados
if all(uploaded_files):
    
    if st.button("▶️ Iniciar Análise"):
        
        progress_bar = st.progress(0, text="Iniciando análise...")
        
        if tipo_analise == "Análise Financeira (R$)":
            coluna_calculo = "Valor líquido"
            formatador = formatar_numero_financeiro
        else:
            coluna_calculo = "Quantidade"
            formatador = formatar_numero_quantidade
        
        selected_centers = st.multiselect(
            "Selecione os centros para análise:",
            options=[1001, 1002, 1003],
            default=[1001, 1002]
        )
        
        try:
            time.sleep(0.5)
            progress_bar.progress(10, text="Carregando arquivos...")
            dfs = []
            for file in uploaded_files:
                if file.name.endswith('.csv'):
                    dfs.append(pd.read_csv(file))
                else:
                    dfs.append(pd.read_excel(file))
        except Exception as e:
            st.error(f"Erro ao ler um dos arquivos. Verifique os formatos. Detalhe: {e}")
            progress_bar.empty()
            st.stop()
            
        time.sleep(0.5)
        progress_bar.progress(30, text="Arquivos carregados. Validando datas...")
        meses, anos = [], []
        for i, df in enumerate(dfs):
            if 'Dt Lanct' not in df.columns:
                st.error(f"Arquivo {i+1} não possui a coluna 'Dt Lanct'.")
                progress_bar.empty()
                st.stop()
            
            df['Dt Lanct'] = pd.to_datetime(df['Dt Lanct'], format='%d.%m.%Y', errors='coerce')
            df.dropna(subset=['Dt Lanct'], inplace=True)
            
            if df.empty:
                st.error(f"Arquivo {i+1} não contém datas válidas na 'Dt Lanct'.")
                progress_bar.empty()
                st.stop()
            
            meses.append(df['Dt Lanct'].dt.month.mode()[0])
            anos.append(df['Dt Lanct'].dt.year.mode()[0])

        if len(set(anos)) > 1:
            st.error(f"Arquivos de anos diferentes detectados: {list(set(anos))}.")
            progress_bar.empty()
            st.stop()
        ano_analise = anos[0]

        quarter_encontrado = next((q for q in MESES_POR_QUARTER.values() if sorted(meses) == sorted(q['meses'])), None)

        if not quarter_encontrado:
            st.error(f"Os meses dos arquivos ({sorted(meses)}) não formam um quarter completo.")
            progress_bar.empty()
            st.stop()

        time.sleep(0.5)
        progress_bar.progress(60, text="Datas validadas. Calculando resultados...")
        meses_anos_str = [f"{df['Dt Lanct'].dt.month.mode()[0]:02d}/{ano_analise}" for df in dfs]
        
        resultados_individuais = [calcular_desempenho(df, selected_centers, coluna_calculo) for df in dfs]
        resultado_soma = calcular_desempenho(pd.concat(dfs, ignore_index=True), selected_centers, coluna_calculo)

        time.sleep(0.5)
        progress_bar.progress(90, text="Análise concluída. Gerando visualização...")

        results_container = st.container()
        with results_container:
            st.markdown("---")
            for resultado, mes_ano in zip(resultados_individuais, sorted(meses_anos_str)):
                st.markdown(f"### Resultado Mensal - {mes_ano}")
                col1, col2, col3 = st.columns(3)
                with col2:
                    exibir_kpi("Total Geral", formatador(resultado["Total Geral"]))
                
                c1, c2, c3 = st.columns(3)
                with c1:
                    exibir_kpi("Total Local", formatador(resultado["Total Local"]))
                    exibir_kpi("% - Local", f"{resultado['% - Local']:.2%}")
                with c2:
                    exibir_kpi("Total Fora", formatador(resultado["Total Fora"]))
                    exibir_kpi("% - Fora", f"{resultado['% - Fora']:.2%}")
                with c3:
                    exibir_kpi("Total Importado", formatador(resultado["Total Importado"]))
                    exibir_kpi("% - Importação", f"{resultado['% - Importação']:.2%}")
                
                c4, c5, c6 = st.columns(3)
                with c4:
                    exibir_kpi("Total Beneficiamento", formatador(resultado["Total Beneficiamento"]))
                    exibir_kpi("% - Beneficiamento", f"{resultado['% - Beneficiamento']:.2%}")
                with c5:
                    exibir_kpi("Total Sucata", formatador(resultado["Total Sucata"]))
                    exibir_kpi("% - Sucata", f"{resultado['% - Sucata']:.2%}")
                with c6:
                    exibir_kpi("Total Nacional Fora", formatador(resultado["Total Nacional Fora"]))
                    exibir_kpi("% - Nacional Fora", f"{resultado['% - Nacional Fora']:.2%}")

            st.markdown("---")
            st.markdown(f"### 📊 Total Consolidado {quarter_encontrado['quarter']} {ano_analise}")
            col_total1, col_total2, col_total3 = st.columns(3)
            with col_total2:
                exibir_kpi("Total Geral", formatador(resultado_soma["Total Geral"]))

            c1, c2, c3 = st.columns(3)
            with c1:
                exibir_kpi("Total Local", formatador(resultado_soma["Total Local"]))
                exibir_kpi("% - Local", f"{resultado_soma['% - Local']:.2%}")
            with c2:
                exibir_kpi("Total Fora", formatador(resultado_soma["Total Fora"]))
                exibir_kpi("% - Fora", f"{resultado_soma['% - Fora']:.2%}")
            with c3:
                exibir_kpi("Total Importado", formatador(resultado_soma["Total Importado"]))
                exibir_kpi("% - Importação", f"{resultado_soma['% - Importação']:.2%}")
                
            c4, c5, c6 = st.columns(3)
            with c4:
                exibir_kpi("Total Beneficiamento", formatador(resultado_soma["Total Beneficiamento"]))
                exibir_kpi("% - Beneficiamento", f"{resultado['% - Beneficiamento']:.2%}")
            with c5:
                exibir_kpi("Total Sucata", formatador(resultado_soma["Total Sucata"]))
                exibir_kpi("% - Sucata", f"{resultado['% - Sucata']:.2%}")
            with c6:
                exibir_kpi("Total Nacional Fora", formatador(resultado_soma["Total Nacional Fora"]))
                exibir_kpi("% - Nacional Fora", f"{resultado['% - Nacional Fora']:.2%}")
            
            st.markdown("---")
            if st.button("📤 Salvar no Google Sheets"):
                with st.spinner("Salvando dados na planilha..."):
                    if atualizar_google_sheets(resultado_soma, quarter_encontrado['quarter'], ano_analise):
                        st.success("Dados salvos com sucesso no Google Sheets!")
        
        time.sleep(0.5)
        progress_bar.progress(100, text="Concluído!")
        time.sleep(1)
        progress_bar.empty()