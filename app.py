"""
Aplicação principal Streamlit para processamento de relatórios trimestrais.
"""
import streamlit as st
import pandas as pd
import time
import os
import io
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

# --- LÓGICA PRINCIPAL DO APP ---

# --- LINHAS COMENTADAS PARA DESATIVAR A SELEÇÃO ---
# st.markdown("### Selecione o tipo de análise:")
# tipo_analise = st.selectbox(
#     "Análise",
#     ("Análise Financeira (R$)", "Análise por Quantidade"),
#     label_visibility="collapsed"
# )
# --- FIM DAS LINHAS COMENTADAS ---

# --- VARIÁVEIS FIXADAS PARA O MODO "QUANTIDADE" ---
# Força a aplicação a rodar sempre a análise por quantidade.
tipo_analise = "Análise por Quantidade"
# --- FIM DAS VARIÁVEIS FIXADAS ---


nome_analise_arquivo = "Financeiro" if tipo_analise == "Análise Financeira (R$)" else "Quantidade"

uploaded_conversion_file = None
# Esta condição agora será sempre verdadeira, mostrando o campo de upload de conversão
if tipo_analise == "Análise por Quantidade":
    st.markdown("### Faça o upload da planilha de conversão de unidades:")
    uploaded_conversion_file = st.file_uploader(
        "Planilha de Conversão",
        type=['xlsx', 'csv'],
        key="conversion_file"
    )

st.markdown("### Faça o upload dos três relatórios para a análise trimestral:")
uploaded_files = [st.file_uploader(f"Relatório do mês {i+1}", type=['xlsx', 'csv'], key=f"file{i}") for i in range(3)]

iniciar_analise = False
if tipo_analise == "Análise Financeira (R$)" and all(uploaded_files):
    iniciar_analise = True
elif tipo_analise == "Análise por Quantidade" and all(uploaded_files) and uploaded_conversion_file:
    iniciar_analise = True

if iniciar_analise:
    if st.button("▶️ Iniciar Análise"):
        progress_bar = st.progress(0, text="Iniciando análise...")
        coluna_calculo = "Valor líquido" if tipo_analise == "Análise Financeira (R$)" else "Quantidade"
        formatador = formatar_numero_financeiro if coluna_calculo == "Valor líquido" else formatar_numero_quantidade
        
        mapa_conversao = None
        if tipo_analise == "Análise por Quantidade":
            try:
                progress_bar.progress(5, text="Carregando arquivo de conversão...")
                mapa_conversao = carregar_e_preparar_conversao(uploaded_conversion_file)
            except Exception as e:
                st.error(f"Erro ao ler o arquivo de conversão: {e}")
                st.stop()
        
        try:
            progress_bar.progress(10, text="Carregando e limpando arquivos...")
            dfs = []
            for file in uploaded_files:
                df = pd.read_csv(file, dtype={'Material': str}) if file.name.endswith('.csv') else pd.read_excel(file, dtype={'Material': str})
                
                if 'Material' in df.columns:
                    df['Material'] = df['Material'].astype(str).str.strip()
                if 'UF' in df.columns:
                    df['UF'] = df['UF'].astype(str).str.strip()

                for col in ['Quantidade', 'Valor líquido']:
                    if col in df.columns:
                        series = df[col].astype(str)
                        series = series.str.replace('R$', '', regex=False).str.strip()
                        is_br_format = series.str.contains(',', na=False)
                        series.loc[is_br_format] = series.loc[is_br_format].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                        series.loc[~is_br_format] = series.loc[~is_br_format].str.replace(',', '', regex=False)
                        df[col] = pd.to_numeric(series, errors='coerce').fillna(0)
                
                dfs.append(df)
            
            df_completo = pd.concat(dfs, ignore_index=True)
            centros_disponiveis = sorted(df_completo['Centro'].dropna().unique().astype(int))
            selected_centers = st.multiselect("Selecione os centros:", options=centros_disponiveis, default=centros_disponiveis)
        except Exception as e:
            st.error(f"Erro ao ler ou limpar os arquivos: {e}"); st.stop()

        if not selected_centers: st.warning("Selecione pelo menos um centro."); st.stop()

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
        
        resultados_individuais = [
            calcular_desempenho(df, selected_centers, coluna_calculo, mapa_conversao=mapa_conversao, debug_contexto=f"Mês {mes_ano}")
            for df, mes_ano in zip(dfs_ordenados, meses_anos_str)
        ]
        
        nome_consolidado = f"Consolidado {quarter_encontrado['quarter']} {ano_analise}"
        resultado_soma = calcular_desempenho(df_completo, selected_centers, coluna_calculo, mapa_conversao=mapa_conversao, debug_contexto=nome_consolidado)
        
        if resultado_soma.get('status') == 'erro_conversao':
            st.error("ERRO DE VALIDAÇÃO: Conversão de Unidades Falhou")
            st.warning("A análise foi interrompida pois os seguintes materiais precisam de uma regra de conversão na sua planilha, mas não foram encontrados. Por favor, adicione-os e tente novamente.")
            df_faltantes = pd.DataFrame(resultado_soma['materiais_faltantes'], columns=['Material', 'Descrição', 'Unidade de Medida'])
            st.dataframe(df_faltantes)
            st.stop()

        progress_bar.progress(90, text="Gerando visualização...")

        with st.container():
            st.markdown("---")
            for resultado, mes_ano in zip(resultados_individuais, meses_anos_str):
                st.markdown(f"### Resultado Mensal - {mes_ano}")
                _, mid_col, _ = st.columns(3);
                with mid_col: exibir_kpi("Total Geral", formatador(resultado["Total Geral"]))
                c1,c2,c3=st.columns(3);
                with c1: exibir_kpi("Total Local", formatador(resultado["Total Local"])); exibir_kpi("% - Local", f"{resultado['% - Local']:.2%}")
                with c2: exibir_kpi("Total Fora", formatador(resultado["Total Fora"])); exibir_kpi("% - Fora", f"{resultado['% - Fora']:.2%}")
                with c3: exibir_kpi("Total Importado", formatador(resultado["Total Importado"])); exibir_kpi("% - Importação", f"{resultado['% - Importação']:.2%}")
                c4,c5,c6=st.columns(3);
                with c4: exibir_kpi("Total Beneficiamento", formatador(resultado["Total Beneficiamento"])); exibir_kpi("% - Beneficiamento", f"{resultado['% - Beneficiamento']:.2%}")
                with c5: exibir_kpi("Total Sucata", formatador(resultado["Total Sucata"])); exibir_kpi("% - Sucata", f"{resultado['% - Sucata']:.2%}")
                with c6: exibir_kpi("Total Nacional Fora", formatador(resultado["Total Nacional Fora"])); exibir_kpi("% - Nacional Fora", f"{resultado['% - Nacional Fora']:.2%}")

            st.markdown("---")
            st.markdown(f"### 📊 Total Consolidado {quarter_encontrado['quarter']} {ano_analise}")
            _, mid_col, _ = st.columns(3)
            with mid_col: exibir_kpi("Total Geral", formatador(resultado_soma["Total Geral"]))
            c1,c2,c3=st.columns(3)
            with c1: exibir_kpi("Total Local", formatador(resultado_soma["Total Local"])); exibir_kpi("% - Local", f"{resultado_soma['% - Local']:.2%}")
            with c2: exibir_kpi("Total Fora", formatador(resultado_soma["Total Fora"])); exibir_kpi("% - Fora", f"{resultado_soma['% - Fora']:.2%}")
            with c3: exibir_kpi("Total Importado", formatador(resultado_soma["Total Importado"])); exibir_kpi("% - Importação", f"{resultado_soma['% - Importação']:.2%}")
            c4,c5,c6=st.columns(3)
            with c4: exibir_kpi("Total Beneficiamento", formatador(resultado_soma["Total Beneficiamento"])); exibir_kpi("% - Beneficiamento", f"{resultado_soma['% - Beneficiamento']:.2%}")
            with c5: exibir_kpi("Total Sucata", formatador(resultado_soma["Total Sucata"])); exibir_kpi("% - Sucata", f"{resultado_soma['% - Sucata']:.2%}")
            with c6: exibir_kpi("Total Nacional Fora", formatador(resultado_soma["Total Nacional Fora"])); exibir_kpi("% - Nacional Fora", f"{resultado_soma['% - Nacional Fora']:.2%}")

            st.markdown("---")
            excel_data = salvar_relatorio_excel(
                resultados_individuais, 
                resultado_soma, 
                meses_anos_str, 
                quarter_encontrado['quarter'], 
                ano_analise,
                nome_analise_arquivo
            )
            st.download_button(
                label="📥 Salvar Relatório em Excel",
                data=excel_data,
                file_name=f"Relatorio_{nome_analise_arquivo}_{quarter_encontrado['quarter']}_{ano_analise}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        progress_bar.progress(100, text="Concluído!")
        time.sleep(1)
        progress_bar.empty()