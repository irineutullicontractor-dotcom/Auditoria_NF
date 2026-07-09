import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Auditoria Interna NF - Produto", layout="wide")

st.title("📊 Auditoria Interna NF - Produto")
st.markdown("""
### Instruções de uso:
1. Carregue o relatório de **NF's** - Puxar relatório do mês vigente.
2. Carregue o relatório de **Credores**.
3. Carregue o relatório do **Painel** - Puxar relatório de no mínimo 90 dias atrás até a data vigente.
4. Carregue o relatório de **Pedidos** - Puxar relatório de no mínimo 90 dias atrás até a data vigente.
5. Carregue o relatório de **Contratos** - Puxar relatório de 01/01/2020 até a data vigente.
""")

# --- UPLOAD DOS FICHEIROS ---
col1, col2 = st.columns(2)
with col1:
    file_nf_prod = st.file_uploader("1. Relatório de NF's - Home / Notas Fiscais / Recepção de NF-e / Relatórios / Notas Fiscais Recebidas.", type=['xlsx'])
    file_forn = st.file_uploader("2. Relatório de Credores - Home / Mais Opções / Apoio / Relatórios / Pessoas / Credores.", type=['xlsx', 'csv'])
    file_painel = st.file_uploader("3. Relatório Painel - Home / Suprimentos / Compras / Painel de Compras (Novo).", type=['xlsx', 'csv'])
with col2:
    file_relacao = st.file_uploader("4. Relatório Pedidos - Home / Suprimentos / Compras / Relatórios / Pedidos de compra / Relação de Pedidos de Compra (Novo).", type=['xlsx', 'csv'])
    file_contrato = st.file_uploader("5. Relatório Contrato - Home / Suprimentos / Contratos e Medições / Relatórios / Contratos / Emissão de Contratos.", type=['xlsx', 'csv'])

# --- FUNÇÕES DE LIMPEZA ---
def limpar_cnpj(v):
    if pd.isna(v): return ""
    num = "".join(filter(str.isdigit, str(v)))
    return num.zfill(14) if len(num) > 11 else num.zfill(11)

def limpar_cod(v):
    if pd.isna(v): return ""
    return str(v).split('.')[0].strip().lstrip('0')

# 🔥 NF PRODUTO → antes da /
def extrair_nf_produto(v):
    if pd.isna(v) or str(v).strip() == "" or str(v).lower() == "nan": return ""
    v = str(v).split('/')[0]
    v = "".join(filter(str.isdigit, v))
    return v.lstrip('0')

# 🔥 NF PAINEL → depois da /
def extrair_nf_painel(v):
    if pd.isna(v) or str(v).strip() == "" or str(v).lower() == "nan": return ""
    v = str(v)
    if '/' in v:
        v = v.split('/')[-1]
    v = "".join(filter(str.isdigit, v))
    return v.lstrip('0')

def estruturar_notas_produtos_interno(file):
    df_bruto = pd.read_excel(file, header=None)
    registros = []
    cnpj_dest = None
    colunas_id = None
    processando = False

    for i, row in df_bruto.iterrows():
        val_a = str(row[0]).strip() if pd.notna(row[0]) else ""
        if "CNPJ do destinatário:" in val_a:
            cnpj_dest = limpar_cnpj(row[3])
            processando = False
            continue
        if val_a == "Emitente":
            colunas_id = [str(c).strip() for c in row.values]
            processando = True
            continue
        if processando and val_a != "" and val_a != "nan":
            registros.append([cnpj_dest] + list(row.values))

    df = pd.DataFrame(registros, columns=['CNPJ Destinatário'] + colunas_id)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed|^nan|None', case=False, na=False)]
    return df.dropna(subset=['Emitente'])

def transformar_credor_limpo(df_bruto):
    for i in range(min(15, len(df_bruto))):
        row_values = [str(x).strip() for x in df_bruto.iloc[i].values]
        if 'Credor' in row_values and 'CNPJ/CPF' in row_values:
            df_header = df_bruto.iloc[i+1:].copy()
            df_header.columns = [str(c).strip() for c in df_bruto.iloc[i].values]
            df_header = df_header.loc[:, df_header.columns.notna() & (df_header.columns != 'nan')]
            def split_safe(val):
                s = str(val).strip()
                return (s.split(" - ")[0], " - ".join(s.split(" - ")[1:])) if " - " in s else ("", s)
            res_split = df_header['Credor'].apply(split_safe)
            df_header['Cód. Fornecedor'] = res_split.apply(lambda x: x[0])
            df_header['Fornecedor'] = res_split.apply(lambda x: x[1])
            return df_header.rename(columns={'CNPJ/CPF': 'CNPJCPF'})
    return df_bruto

# --- PROCESSAMENTO ---
if st.button("🚀 Iniciar Auditoria"):
    if all([file_nf_prod, file_forn, file_painel, file_relacao, file_contrato]):
        
        df_nf = estruturar_notas_produtos_interno(file_nf_prod)
        df_forn = transformar_credor_limpo(pd.read_excel(file_forn, header=None))
        df_painel = pd.read_excel(file_painel)
        df_relacao = pd.read_excel(file_relacao)
        df_bruto_ct = pd.read_excel(file_contrato, header=None)

        # Padronização
        df_forn['CNPJCPF'] = df_forn['CNPJCPF'].apply(limpar_cnpj)
        df_nf['CNPJ emitente'] = df_nf['CNPJ emitente'].apply(limpar_cnpj)

        # 🔥 CORREÇÃO AQUI
        df_nf['nf_limpa'] = df_nf['Núm/Série'].apply(extrair_nf_produto)
        
        # Chave
        df_nf['chave_unica'] = df_nf.apply(lambda r: r['CNPJ emitente'] + "_" + r['nf_limpa'] if r['nf_limpa'] != "" else "SEM_NF_" + str(r.name), axis=1)

        # --- PAINEL ---
        df_painel['nf_ref_limpa'] = df_painel['N° da Nota fiscal'].apply(extrair_nf_painel)
        df_painel['Fornecedor_UP'] = df_painel['Fornecedor'].astype(str).str.strip().str.upper()
        df_forn['Credor_UP'] = df_forn['Credor'].astype(str).str.strip().str.upper()
        
        painel_com_cnpj = pd.merge(df_painel, df_forn[['Credor_UP', 'CNPJCPF']], left_on='Fornecedor_UP', right_on='Credor_UP', how='left')
        painel_com_cnpj['CNPJCPF'] = painel_com_cnpj['CNPJCPF'].apply(limpar_cnpj)
        painel_com_cnpj['chave_p'] = painel_com_cnpj['CNPJCPF'] + "_" + painel_com_cnpj['nf_ref_limpa']
        
        chaves_lancadas = set(painel_com_cnpj[painel_com_cnpj['nf_ref_limpa'] != ""]['chave_p'].unique())
        cnpjs_no_painel = set(painel_com_cnpj['CNPJCPF'].unique())

        # --- ABA 1: PAINEL ---
        resumo_painel = pd.merge(df_nf, painel_com_cnpj[['chave_p', 'N° da Nota fiscal']].drop_duplicates('chave_p'), left_on='chave_unica', right_on='chave_p', how='left')
        
        def status_painel(r):
            if r['chave_unica'] in chaves_lancadas: return "✅ NF Lançada"
            if r['CNPJ emitente'] in cnpjs_no_painel: return "⚠️ Para Verificação"
            return "❌ Sem Histórico"
        
        resumo_painel['Status'] = resumo_painel.apply(status_painel, axis=1)

        # --- PEDIDOS ---
        df_relacao['Cód. fornecedor'] = df_relacao['Cód. fornecedor'].apply(limpar_cod)
        rel_com_cnpj = pd.merge(df_relacao, df_forn[['Cód. Fornecedor', 'CNPJCPF']], left_on='Cód. fornecedor', right_on='Cód. Fornecedor', how='left')
        rel_com_cnpj['CNPJCPF'] = rel_com_cnpj['CNPJCPF'].apply(limpar_cnpj)
        cnpjs_com_pedido = set(rel_com_cnpj['CNPJCPF'].unique())
        
        peds_agrupados = rel_com_cnpj.dropna(subset=['Nº do pedido']).groupby('CNPJCPF')['Nº do pedido'].apply(lambda x: ", ".join(sorted(set(x.astype(str).unique())))).reset_index()

        resumo_pedidos = pd.merge(resumo_painel, peds_agrupados, left_on='CNPJ emitente', right_on='CNPJCPF', how='left')

        def status_pedidos(r):
            if r['Status'] == "✅ NF Lançada": return "✅ Resolvido Painel"
            if r['CNPJ emitente'] in cnpjs_com_pedido or r['Status'] == "⚠️ Para Verificação": 
                return "⚠️ Para Verificação"
            return "❌ Sem Histórico"
        
        resumo_pedidos['Status_Ped'] = resumo_pedidos.apply(status_pedidos, axis=1)

        # --- CONTRATO ---
        registros_ct = []
        item_atual = {'Contrato': None, 'CNPJ': None}
        for i in range(len(df_bruto_ct)):
            l = df_bruto_ct.iloc[i]
            col_a = str(l[0]).strip() if pd.notna(l[0]) else ""
            if col_a == "Contrato": item_atual['Contrato'] = str(l[3]).strip()
            elif col_a == "CNPJ" and item_atual['Contrato']:
                item_atual['CNPJ'] = limpar_cnpj(l[3])
                registros_ct.append(item_atual.copy())
        
        cts_agrupados = pd.DataFrame(registros_ct).groupby('CNPJ')['Contrato'].apply(lambda x: ", ".join(sorted(set(x.astype(str).unique())))).reset_index() if registros_ct else pd.DataFrame(columns=['CNPJ', 'Contrato'])

        resumo_contratos = pd.merge(resumo_pedidos, cts_agrupados, left_on='CNPJ emitente', right_on='CNPJ', how='left')

        def status_ct(r):
            if r['Status_Ped'] == "✅ Resolvido Painel": return "✅ Resolvido Painel"
            if pd.notna(r['Contrato']) and str(r['Contrato']).strip() != "": return "📄 Vínculo Contratual"
            return r['Status_Ped']
        
        resumo_contratos['Status_CT'] = resumo_contratos.apply(status_ct, axis=1)

        # --- SAÍDA FINAL (SEU PADRÃO ORIGINAL) ---
        cols_base = ['Núm/Série', 'CNPJ emitente', 'Emitente', 'Emissão', 'Valor']
        cols_extra = ['CNPJ Destinatário', 'Destinatário']
        
        aba1 = resumo_painel[cols_base + ['N° da Nota fiscal', 'Status'] + cols_extra]
        aba2 = resumo_pedidos[cols_base + ['N° da Nota fiscal', 'Nº do pedido', 'Status_Ped'] + cols_extra].rename(columns={'Status_Ped': 'Status', 'Nº do pedido': 'Pedido'})
        aba3 = resumo_contratos[cols_base + ['N° da Nota fiscal', 'Nº do pedido', 'Contrato', 'Status_CT'] + cols_extra].rename(columns={'Status_CT': 'Status', 'Nº do pedido': 'Pedido'})

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            aba1.to_excel(writer, sheet_name='1. PAINEL', index=False)
            aba2.to_excel(writer, sheet_name='2. PEDIDOS', index=False)
            aba3.to_excel(writer, sheet_name='3. CONTRATO', index=False)
        
        st.success("Tudo pronto! Relatório de Auditoria gerado com sucesso.")
        st.download_button("📥 Baixar Auditoria", output.getvalue(), "AUDITORIA_NF_PRODUTO.xlsx")
