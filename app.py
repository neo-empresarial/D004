import streamlit as st
import pandas as pd
# Certifique-se de que os arquivos style.py e utils.py estão na mesma pasta
from style import get_css_block, get_header_html
from utils import processar_dados_geral, gerar_excel_socioeconomico

st.set_page_config(page_title="Relatório Socioeconômico", page_icon="📊", layout="wide")
st.markdown(get_css_block(), unsafe_allow_html=True)
st.markdown(get_header_html(), unsafe_allow_html=True)

st.title("Gerador de Relatório Socioeconômico")
st.markdown("---")

col_periodo, col_fator, _ = st.columns([1, 1, 2])
with col_periodo:
    qtd_meses = st.number_input("Qtd. Meses", min_value=1, max_value=12, value=1)
with col_fator:
    fator_pis_cofins = st.number_input("Fator Conversão (PIS/COFINS)", value=0.9075, format="%.4f", step=0.0001)

uploaded_files = []
cols = st.columns(3)
for i in range(qtd_meses):
    with cols[i % 3]:
        file = st.file_uploader(f"Mês {i+1}", type=['xlsx', 'csv'], key=f"file_{i}")
        if file: uploaded_files.append(file)

if len(uploaded_files) == qtd_meses and st.button("▶️ Processar Dados", type="primary"):
    with st.spinner("Consolidando dados..."):
        try:
            resultados = processar_dados_geral(uploaded_files, fator_reducao=fator_pis_cofins)
            nome_arquivo = resultados['nome_arquivo']
            
            # --- FRENTE 1: Fornecedores ---
            st.markdown("### 1. Fornecedores (Faturamento)")
            st.markdown("Agora incluindo **UF** e **Classificação** (Local/Importado/Fora).")
            
            df_forn = resultados['df_fornecedores']
            # Exibe colunas principais primeiro
            cols_show = ['Nome Cliente/Forn', 'UF', 'Classificação', 'Faturamento Total']
            # Garante que as colunas existem antes de tentar exibir
            cols_show = [c for c in cols_show if c in df_forn.columns]
            
            st.dataframe(
                df_forn[cols_show].head(10).style.format({"Faturamento Total": "R$ {:,.2f}"}), 
                use_container_width=True
            )
            
            with st.expander("Ver lista completa de fornecedores"):
                st.dataframe(df_forn)
            
            st.markdown("---")
            
            # --- FRENTE 2: Socioeconômico Detalhado ---
            st.markdown(f"### 2. Indicadores Socioeconômicos - {nome_arquivo.replace('Relatorio_Socioeconomico_', '').replace('_', '/')}")
            
            df_display = resultados['df_resumo_socio'].copy()
            
            # Formatação para exibição na tela (Visual apenas)
            format_dict = {
                "Financeiro (R$)": "R$ {:,.2f}",
                "Quantidade (KG)": "{:,.2f}",
                "% Fin (vs MP)": "{:.2%}",
                "% Qtd (vs MP)": "{:.2%}",
                "% Fin (vs Geral)": "{:.2%}"
            }
            
            # Removemos a coluna auxiliar 'Tipo_Linha' da visualização, mas usamos para colorir se necessário
            # Aqui apenas mostramos a tabela limpa
            cols_view = [c for c in df_display.columns if c != 'Tipo_Linha']
            
            st.dataframe(
                df_display[cols_view].style.format(format_dict).apply(
                    lambda x: ['font-weight: bold' if "Total" in str(x['Categoria']) else '' for i in x], axis=1
                ),
                use_container_width=True,
                height=500 # Aumentei a altura pois agora tem mais linhas
            )

            # Dados de Base
            with st.expander("🔍 Ver Dados de Base (Denominadores)"):
                meta = resultados['meta_dados']
                c1, c2, c3 = st.columns(3)
                c1.metric("Total MP (R$)", f"R$ {meta['MP_Fin']:,.2f}")
                c2.metric("Total MP (KG)", f"{meta['MP_Kg']:,.2f}")
                c3.metric("Total Geral Líquido (R$)", f"R$ {meta['Geral_Fin']:,.2f}")

            # Download
            st.markdown("---")
            excel_data = gerar_excel_socioeconomico(resultados)
            st.download_button(
                label=f"📥 Baixar Relatório Excel ({nome_arquivo}.xlsx)",
                data=excel_data,
                file_name=f"{nome_arquivo}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro: {str(e)}")
            import traceback; traceback.print_exc()