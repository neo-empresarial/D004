"""
Aplicação principal Streamlit para processamento de relatórios trimestrais.
Permite a análise por valor financeiro ou por quantidade.
"""
import streamlit as st
import pandas as pd
import time
import os
from style import get_css_block, get_header_html
from utils import (
    carregar_e_preparar_conversao,
    calcular_desempenho,
    formatar_numero_financeiro,
    formatar_numero_quantidade,
    exibir_kpi,
    salvar_relatorio_excel,
    MESES_POR_QUARTER
)

# --- Configuração da Página ---
st.set_page_config(page_title="Processador de Relatórios Docol", page_icon="📊", layout="wide")
st.markdown(get_css_block(), unsafe_allow_html=True)
st.markdown(get_header_html(), unsafe_allow_html=True)
pd.set_option("styler.render.max_elements", 1_000_000)

# --- Carregamento do Arquivo de Conversão ---
try:
    diretorio_atual = os.path.dirname(__file__)
    CAMINHO_ARQUIVO_CONVERSAO = os.path.join(diretorio_atual, "Input de conversão", "Dados de conversão.xlsx")
except NameError:
    CAMINHO_ARQUIVO_CONVERSAO = os.path.join(os.getcwd(), "Input de conversão", "Dados de conversão.xlsx")

@st.cache_data
def carregar_dados_conversao():
    return carregar_e_preparar_conversao(CAMINHO_ARQUIVO_CONVERSAO)

df_conversao, mapa_moda = carregar_dados_conversao()

# --- LÓGICA PRINCIPAL DO APP ---
if df_conversao is None:
    st.error("Erro Crítico: Não foi possível carregar o arquivo de conversão.")
else:
    st.markdown("### Selecione o tipo de análise:")
    tipo_analise = st.selectbox("Análise", ("Análise Financeira (R$)", "Análise por Quantidade"), label_visibility="collapsed")

    st.markdown("### Faça o upload dos três relatórios para a análise trimestral:")
    uploaded_files = [st.file_uploader(f"Relatório do mês {i+1}", type=['xlsx', 'csv'], key=f"file{i}") for i in range(3)]

    if all(uploaded_files):
        if st.button("▶️ Iniciar Análise"):
            progress_bar = st.progress(0, text="Iniciando análise...")
            coluna_calculo = "Valor líquido" if tipo_analise == "Análise Financeira (R$)" else "Quantidade"
            formatador = formatar_numero_financeiro if coluna_calculo == "Valor líquido" else formatar_numero_quantidade
            
            # --- ADICIONADO: Define o nome do arquivo com base na análise ---
            nome_analise_arquivo = "Financeiro" if tipo_analise == "Análise Financeira (R$)" else "Quantidade"
            
            try:
                progress_bar.progress(10, text="Carregando arquivos...")
                dfs = [pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file) for file in uploaded_files]
                df_completo = pd.concat(dfs, ignore_index=True)
                centros_disponiveis = sorted(df_completo['Centro'].dropna().unique().astype(int))
                selected_centers = st.multiselect("Selecione os centros:", options=centros_disponiveis, default=centros_disponiveis)
            except Exception as e:
                st.error(f"Erro ao ler os arquivos: {e}"); st.stop()

            if not selected_centers:
                st.warning("Selecione pelo menos um centro."); st.stop()

            progress_bar.progress(30, text="Validando datas...")
            meses, anos = [], []
            for i, df in enumerate(dfs):
                if 'Dt Lanct' not in df.columns: st.error(f"Arquivo {i+1} não tem 'Dt Lanct'."); st.stop()
                df['Dt Lanct'] = pd.to_datetime(df['Dt Lanct'], format='%d.%m.%Y', errors='coerce')
                df.dropna(subset=['Dt Lanct'], inplace=True)
                if df.empty: st.error(f"Arquivo {i+1} não tem datas válidas."); st.stop()
                meses.append(df['Dt Lanct'].dt.month.mode()[0]); anos.append(df['Dt Lanct'].dt.year.mode()[0])
            
            if len(set(anos)) > 1: st.error(f"Arquivos de anos diferentes: {list(set(anos))}."); st.stop()
            ano_analise = anos[0]
            quarter_encontrado = next((q for q in MESES_POR_QUARTER.values() if sorted(meses) == q['meses']), None)
            if not quarter_encontrado: st.error(f"Os meses ({sorted(meses)}) não formam um quarter."); st.stop()
            
            progress_bar.progress(60, text="Calculando resultados...")
            dfs_ordenados = sorted(dfs, key=lambda d: d['Dt Lanct'].dt.month.mode()[0])
            meses_anos_str = [f"{df['Dt Lanct'].dt.month.mode()[0]:02d}/{ano_analise}" for df in dfs_ordenados]
            resultados_individuais = [calcular_desempenho(df, selected_centers, coluna_calculo, df_conversao, mapa_moda) for df in dfs_ordenados]
            resultado_soma = calcular_desempenho(df_completo, selected_centers, coluna_calculo, df_conversao, mapa_moda)

            progress_bar.progress(90, text="Gerando visualização...")
            
            with st.container():
                st.markdown("---")
                # Bloco para exibir resultados mensais
                for resultado, mes_ano in zip(resultados_individuais, meses_anos_str):
                    st.markdown(f"### Resultado Mensal - {mes_ano}")
                    _, mid_col, _ = st.columns(3)
                    with mid_col:
                        exibir_kpi("Total Geral", formatador(resultado["Total Geral"]))

                    c1,c2,c3=st.columns(3)
                    with c1: exibir_kpi("Total Local", formatador(resultado["Total Local"])); exibir_kpi("% - Local", f"{resultado['% - Local']:.2%}")
                    with c2: exibir_kpi("Total Fora", formatador(resultado["Total Fora"])); exibir_kpi("% - Fora", f"{resultado['% - Fora']:.2%}")
                    with c3: exibir_kpi("Total Importado", formatador(resultado["Total Importado"])); exibir_kpi("% - Importação", f"{resultado['% - Importação']:.2%}")

                    c4,c5,c6=st.columns(3)
                    with c4: exibir_kpi("Total Beneficiamento", formatador(resultado["Total Beneficiamento"])); exibir_kpi("% - Beneficiamento", f"{resultado['% - Beneficiamento']:.2%}")
                    with c5: exibir_kpi("Total Sucata", formatador(resultado["Total Sucata"])); exibir_kpi("% - Sucata", f"{resultado['% - Sucata']:.2%}")
                    with c6: exibir_kpi("Total Nacional Fora", formatador(resultado["Total Nacional Fora"])); exibir_kpi("% - Nacional Fora", f"{resultado['% - Nacional Fora']:.2%}")

                st.markdown("---")
                # Bloco para exibir resultado consolidado
                st.markdown(f"### 📊 Total Consolidado {quarter_encontrado['quarter']} {ano_analise}")
                _, mid_col, _ = st.columns(3)
                with mid_col:
                    exibir_kpi("Total Geral", formatador(resultado_soma["Total Geral"]))

                c1,c2,c3=st.columns(3)
                with c1: exibir_kpi("Total Local", formatador(resultado_soma["Total Local"])); exibir_kpi("% - Local", f"{resultado_soma['% - Local']:.2%}")
                with c2: exibir_kpi("Total Fora", formatador(resultado_soma["Total Fora"])); exibir_kpi("% - Fora", f"{resultado_soma['% - Fora']:.2%}")
                with c3: exibir_kpi("Total Importado", formatador(resultado_soma["Total Importado"])); exibir_kpi("% - Importação", f"{resultado_soma['% - Importação']:.2%}")

                c4,c5,c6=st.columns(3)
                with c4: exibir_kpi("Total Beneficiamento", formatador(resultado_soma["Total Beneficiamento"])); exibir_kpi("% - Beneficiamento", f"{resultado_soma['% - Beneficiamento']:.2%}")
                with c5: exibir_kpi("Total Sucata", formatador(resultado_soma["Total Sucata"])); exibir_kpi("% - Sucata", f"{resultado_soma['% - Sucata']:.2%}")
                with c6: exibir_kpi("Total Nacional Fora", formatador(resultado_soma["Total Nacional Fora"])); exibir_kpi("% - Nacional Fora", f"{resultado_soma['% - Nacional Fora']:.2%}")

                st.markdown("---")
                
                # Bloco de Download do Excel
                excel_data = salvar_relatorio_excel(
                    resultados_individuais, 
                    resultado_soma, 
                    meses_anos_str, 
                    quarter_encontrado['quarter'], 
                    ano_analise
                )
                
                st.download_button(
                    label="📥 Salvar Relatório em Excel",
                    data=excel_data,
                    # --- ALTERADO: Adiciona o tipo de análise ao nome do arquivo ---
                    file_name=f"Relatorio_{nome_analise_arquivo}_{quarter_encontrado['quarter']}_{ano_analise}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            progress_bar.progress(100, text="Concluído!")
            time.sleep(1)
            progress_bar.empty()
