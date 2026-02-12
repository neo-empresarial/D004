import pandas as pd
import numpy as np
import io

# --- Configurações e Constantes ---
UNIDADES_PESO_KG = ['kg', 'quilograma', 'kilograma']
UNIDADES_PESO_TON = ['t', 'ton', 'tonelada', 'to'] # Multiplica por 1000

PREFIXOS_SUCATA = ('10028330000', '10032677000', '10002709000', '10001099000', '10001103000')

# Itens que NÃO sofrem redução de PIS/COFINS (busca por substring "is in")
MATERIAIS_EXCECAO_REDUCAO = ('2833000', '3267700')

FORNECEDORES_BENEFICIAMENTO = [1048374, 1028618]
CFOP_BENEFICIAMENTO = '2124AA'
CFOPS_FRETE = ['1352AA', '2352AA']

# Mapa de meses para nomear arquivo
MAPA_MESES = {
    1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun',
    7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
}

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def formatar_numero(valor):
    return f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def limpar_valor_monetario(series):
    series = series.astype(str).str.replace('R$', '', regex=False).str.strip()
    is_br = series.str.contains(',', na=False)
    series_br = series[is_br].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    series_us = series[~is_br].str.replace(',', '', regex=False)
    series_final = series.copy()
    series_final.update(series_br)
    series_final.update(series_us)
    return pd.to_numeric(series_final, errors='coerce').fillna(0)

def definir_origem(uf):
    uf = str(uf).strip().upper()
    if uf in ['SC', 'PR']:
        return 'Local'
    elif uf == 'EX':
        return 'Importado'
    else:
        return 'Fora'

def processar_dados_geral(files, fator_reducao=0.9075):
    dfs = []
    
    # --- 1. Leitura e Consolidação ---
    for file in files:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file, dtype={'Material': str, 'CFOP': str, 'Cliente/Fornec': str, 'UF': str})
        else:
            df = pd.read_excel(file, dtype={'Material': str, 'CFOP': str, 'Cliente/Fornec': str, 'UF': str})
            
        # Tratamento de Datas para o nome do arquivo
        if 'Dt Lanct' in df.columns:
            df['Dt Lanct'] = pd.to_datetime(df['Dt Lanct'], format='%d.%m.%Y', errors='coerce')
        
        # Normalização básica
        if 'Material' in df.columns:
            df['Material'] = df['Material'].astype(str).str.strip().str.lstrip('0')
        if 'CFOP' in df.columns:
            df['CFOP'] = df['CFOP'].astype(str).str.replace('.', '', regex=False).str.strip()
        if 'Valor líquido' in df.columns:
            df['Valor líquido'] = limpar_valor_monetario(df['Valor líquido'])
        if 'Quantidade' in df.columns:
            df['Quantidade'] = limpar_valor_monetario(df['Quantidade'])
        if 'UF' in df.columns:
            df['UF'] = df['UF'].astype(str).str.strip().str.upper().fillna('')
        else:
            df['UF'] = '' # Garante que a coluna existe
            
        dfs.append(df)
    
    df_completo = pd.concat(dfs, ignore_index=True)

    # Aplica a classificação de origem
    df_completo['Origem_Classificacao'] = df_completo['UF'].apply(definir_origem)

    # --- 2. Identificação do Período (Nome do Arquivo) ---
    meses_presentes = sorted(df_completo['Dt Lanct'].dt.month.dropna().unique().astype(int))
    anos_presentes = df_completo['Dt Lanct'].dt.year.dropna().unique().astype(int)
    
    nome_meses = "_".join([MAPA_MESES.get(m, str(m)) for m in meses_presentes])
    nome_ano = str(anos_presentes[0]) if len(anos_presentes) > 0 else "Ano"
    nome_arquivo_sugerido = f"Relatorio_Socioeconomico_{nome_meses}_{nome_ano}"

    # --- 3. Filtro Global (Remove Frete) ---
    df_completo = df_completo[~df_completo['CFOP'].isin(CFOPS_FRETE)].copy()

    # --- 3.1 Aplicação do Fator PIS/COFINS (NOVA LÓGICA) ---
    # Cria regex para buscar '2833000' ou '3267700' dentro do código do material
    pat_excecao = '|'.join(MATERIAIS_EXCECAO_REDUCAO)
    mask_excecao = df_completo['Material'].astype(str).str.contains(pat_excecao, na=False)
    
    # Aplica o fator apenas onde NÃO for exceção (reduz o valor líquido)
    df_completo.loc[~mask_excecao, 'Valor líquido'] = df_completo.loc[~mask_excecao, 'Valor líquido'] * fator_reducao

    # --- 4. Frente Fornecedores ---
    # Adicionado UF e Origem_Classificacao ao agrupamento
    cols_identificacao = ['Nome Cliente/Forn', 'Cliente/Fornec', 'CPF/CNPJ', 'UF', 'Origem_Classificacao']
    cols_existentes = [c for c in cols_identificacao if c in df_completo.columns]
    
    if cols_existentes:
        df_fornecedores = df_completo.groupby(cols_existentes)['Valor líquido'].sum().reset_index()
        df_fornecedores.rename(columns={'Valor líquido': 'Faturamento Total', 'Origem_Classificacao': 'Classificação'}, inplace=True)
        df_fornecedores = df_fornecedores.sort_values('Faturamento Total', ascending=False)
    else:
        df_fornecedores = pd.DataFrame()

    # --- 5. Lógica de Categorização ---
    
    # A) Sucata
    mask_sucata = df_completo['Material'].str.startswith(PREFIXOS_SUCATA, na=False)
    
    # B) Beneficiamento
    cliente_num = pd.to_numeric(df_completo['Cliente/Fornec'], errors='coerce')
    mask_benef = (cliente_num.isin(FORNECEDORES_BENEFICIAMENTO)) & (df_completo['CFOP'] == CFOP_BENEFICIAMENTO)
    
    # C) Matéria Prima (Centro 1001 + Começa com 10 + Unidade Peso)
    df_completo['Unidade_Normalizada'] = df_completo['Unidade de medida'].astype(str).str.lower().str.strip()
    unidades_mp = UNIDADES_PESO_KG + UNIDADES_PESO_TON
    
    mask_centro_1001 = df_completo['Centro'].astype(str).str.contains('1001', na=False)
    mask_material_10 = df_completo['Material'].str.startswith('10', na=False)
    mask_unidade_peso = df_completo['Unidade_Normalizada'].isin(unidades_mp)
    mask_mp = mask_centro_1001 & mask_material_10 & mask_unidade_peso

    # --- 6. Cálculos de Peso (KG) ---
    df_completo['Peso_KG'] = 0.0
    mask_kg = df_completo['Unidade_Normalizada'].isin(UNIDADES_PESO_KG)
    mask_ton = df_completo['Unidade_Normalizada'].isin(UNIDADES_PESO_TON)
    
    df_completo.loc[mask_kg, 'Peso_KG'] = df_completo.loc[mask_kg, 'Quantidade']
    df_completo.loc[mask_ton, 'Peso_KG'] = df_completo.loc[mask_ton, 'Quantidade'] * 1000.0

    # --- 7. Agregação dos KPI's ---
    
    # Denominadores (Bases)
    total_mp_fin = df_completo.loc[mask_mp, 'Valor líquido'].sum()
    total_mp_kg = df_completo.loc[mask_mp, 'Peso_KG'].sum()
    total_geral_fin = df_completo['Valor líquido'].sum()

    lista_resultados = []

    def adicionar_linhas_categoria(mask_principal, nome_categoria_base):
        # 1. Linha Total da Categoria
        val_fin_total = df_completo.loc[mask_principal, 'Valor líquido'].sum()
        val_kg_total = df_completo.loc[mask_principal, 'Peso_KG'].sum()
        
        lista_resultados.append({
            "Categoria": nome_categoria_base + " (Total)",
            "Financeiro (R$)": val_fin_total,
            "Quantidade (KG)": val_kg_total,
            "% Fin (vs MP)": (val_fin_total / total_mp_fin) if total_mp_fin else 0,
            "% Qtd (vs MP)": (val_kg_total / total_mp_kg) if total_mp_kg else 0,
            "% Fin (vs Geral)": (val_fin_total / total_geral_fin) if total_geral_fin else 0,
            "Tipo_Linha": "Total"
        })
        
        # 2. Granularização (Local, Importado, Fora)
        for origem in ['Local', 'Importado', 'Fora']:
            mask_origem = mask_principal & (df_completo['Origem_Classificacao'] == origem)
            
            val_fin = df_completo.loc[mask_origem, 'Valor líquido'].sum()
            val_kg = df_completo.loc[mask_origem, 'Peso_KG'].sum()
            
            lista_resultados.append({
                "Categoria": f"   ↳ {origem}", # Indentação visual
                "Financeiro (R$)": val_fin,
                "Quantidade (KG)": val_kg,
                "% Fin (vs MP)": (val_fin / total_mp_fin) if total_mp_fin else 0,
                "% Qtd (vs MP)": (val_kg / total_mp_kg) if total_mp_kg else 0,
                "% Fin (vs Geral)": (val_fin / total_geral_fin) if total_geral_fin else 0,
                "Tipo_Linha": "Detalhe"
            })

    # Gerar linhas para Beneficiamento
    adicionar_linhas_categoria(mask_benef, "Beneficiamento")
    
    # Gerar linhas para Sucata
    adicionar_linhas_categoria(mask_sucata, "Sucata")
    
    # Gerar linhas para TOTAL SOCIOECONÔMICO
    mask_total_socio = mask_benef | mask_sucata
    adicionar_linhas_categoria(mask_total_socio, "TOTAL SOCIOECONÔMICO")

    df_resumo_socio = pd.DataFrame(lista_resultados)
    
    # Metadados para exibição
    meta_dados = {
        "MP_Fin": total_mp_fin,
        "MP_Kg": total_mp_kg,
        "Geral_Fin": total_geral_fin
    }

    return {
        'df_fornecedores': df_fornecedores,
        'df_resumo_socio': df_resumo_socio,
        'meta_dados': meta_dados,
        'nome_arquivo': nome_arquivo_sugerido
    }

def gerar_excel_socioeconomico(resultados):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        
        # Formatos
        fmt_money = wb.add_format({'num_format': 'R$ #,##0.00'})
        fmt_kg = wb.add_format({'num_format': '#,##0.00 "kg"'})
        fmt_pct = wb.add_format({'num_format': '0.00%'})
        fmt_bold = wb.add_format({'bold': True})
        fmt_indent = wb.add_format({'indent': 2}) # Para as linhas filhas
        fmt_header = wb.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})

        # --- ABA 1: RESUMO SOCIOECONÔMICO ---
        df_socio = resultados['df_resumo_socio'].drop(columns=['Tipo_Linha']) # Remove coluna aux
        sheet_name = 'Resumo Socioeconômico'
        df_socio.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
        ws = writer.sheets[sheet_name]
        
        # Cabeçalho customizado
        ws.write(0, 0, "Detalhamento Granularizado (Local x Importado x Fora)", fmt_bold)
        
        # Formatação das Colunas
        ws.set_column('A:A', 35) # Categoria
        ws.set_column('B:B', 20, fmt_money) # Financeiro
        ws.set_column('C:C', 20, fmt_kg) # Quantidade
        ws.set_column('D:F', 18, fmt_pct) # Percentuais

        # Adicionar Tabela de Dados de Base (Denominadores) ao lado ou abaixo
        row_base = len(df_socio) + 4
        ws.write(row_base, 0, "DADOS DE BASE (DENOMINADORES)", fmt_bold)
        ws.write(row_base+1, 0, "Total Fin. Matéria Prima")
        ws.write(row_base+1, 1, resultados['meta_dados']['MP_Fin'], fmt_money)
        
        ws.write(row_base+2, 0, "Total Kg Matéria Prima")
        ws.write(row_base+2, 1, resultados['meta_dados']['MP_Kg'], fmt_kg)
        
        ws.write(row_base+3, 0, "Total Fin. Geral (s/ Frete)")
        ws.write(row_base+3, 1, resultados['meta_dados']['Geral_Fin'], fmt_money)

        # --- ABA 2: FORNECEDORES ---
        if not resultados['df_fornecedores'].empty:
            resultados['df_fornecedores'].to_excel(writer, sheet_name='Fornecedores', index=False)
            ws_forn = writer.sheets['Fornecedores']
            ws_forn.set_column('A:A', 40) # Nome
            ws_forn.set_column('D:D', 5)  # UF
            ws_forn.set_column('E:E', 15) # Classificacao
            ws_forn.set_column('F:F', 20, fmt_money) # Valor

    return output.getvalue()