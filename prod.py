import streamlit as st
import pandas as pd
import datetime
import io

st.set_page_config(page_title="Auditoria Interna NF - Produto", layout="wide")

st.title("📊 Auditoria Interna NF - Produto")
st.markdown("""
### Regra de Baixas Globais:
- **Baixa Automática:** Cruzamento de NFs lançadas no **Painel** + **Relatório de Títulos/Financeiro** (Todas as Obras).
- **Status Pendentes e Pedidos:** Filtrado **exclusivamente pela SUA OBRA**.
- **Aba 4. EM ABERTO:** Consolidação de pedidos da sua obra que ainda não possuem Nota Fiscal/Entrada vinculada (com Fornecedor e Dias em Aberto).
""")

codigo_obra_usuario = st.text_input("📍 Informe o código numérico da sua obra (Ex: 2):", value="").strip()

col1, col2 = st.columns(2)
with col1:
    file_nf_prod = st.file_uploader("1. Relatório de NF's", type=['xlsx'])
    file_forn = st.file_uploader("2. Relatório de Credores", type=['xlsx', 'csv'])
    file_painel = st.file_uploader("3. Relatório Painel (Todas as obras)", type=['xlsx', 'csv'])
with col2:
    file_relacao = st.file_uploader("4. Relatório Pedidos (Todas as obras)", type=['xlsx', 'csv'])
    file_contrato = st.file_uploader("5. Relatório Contrato", type=['xlsx', 'csv'])
    file_titulo = st.file_uploader("6. Relatório Título (Todas as obras)", type=['xlsx'])

def limpar_cnpj(v):
    if pd.isna(v): return ""
    num = "".join(filter(str.isdigit, str(v)))
    return num.zfill(14) if len(num) > 11 else num.zfill(11)

def limpar_cod(v):
    if pd.isna(v): return ""
    return str(v).split('.')[0].strip().lstrip('0')

def extrair_nf_produto(v):
    if pd.isna(v) or str(v).strip() == "" or str(v).lower() == "nan": return ""
    v = str(v).split('/')[0]
    v = "".join(filter(str.isdigit, v))
    return v.lstrip('0')

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

def estruturar_titulo_limpo(file):
    df_bruto = pd.read_excel(file, header=None)
    inicio_dados = None
    for i, row in df_bruto.iterrows():
        if str(row[0]).strip().lower() == "item":
            inicio_dados = i
            break
    if inicio_dados is None:
        return pd.DataFrame()
    df = pd.read_excel(file, skiprows=inicio_dados)
    df.columns = [str(c).strip() for c in df.columns]
    return df

if st.button("🚀 Iniciar Auditoria"):
    if not codigo_obra_usuario:
        st.error("⚠️ Por favor, digite o código numérico da sua obra antes de iniciar.")
        st.stop()

    if all([file_nf_prod, file_forn, file_painel, file_relacao, file_contrato, file_titulo]):
        cod_obra_alvo = limpar_cod(codigo_obra_usuario)
        hoje = pd.to_datetime(datetime.date.today())

        df_nf = estruturar_notas_produtos_interno(file_nf_prod)
        df_forn = transformar_credor_limpo(pd.read_excel(file_forn, header=None))
        
        df_painel_raw = pd.read_excel(file_painel) if file_painel.name.endswith('xlsx') else pd.read_csv(file_painel)
        df_relacao_raw = pd.read_excel(file_relacao) if file_relacao.name.endswith('xlsx') else pd.read_csv(file_relacao)
        df_bruto_ct = pd.read_excel(file_contrato, header=None)
        df_titulos_raw = estruturar_titulo_limpo(file_titulo)

        df_forn['CNPJCPF'] = df_forn['CNPJCPF'].apply(limpar_cnpj)
        df_forn['Credor_UP'] = df_forn['Credor'].astype(str).str.strip().str.upper()
        df_nf['CNPJ emitente'] = df_nf['CNPJ emitente'].apply(limpar_cnpj)

        df_nf['nf_limpa'] = df_nf['Núm/Série'].apply(extrair_nf_produto)
        df_nf['chave_unica'] = df_nf.apply(lambda r: r['CNPJ emitente'] + "_" + r['nf_limpa'] if r['nf_limpa'] != "" else "SEM_NF_" + str(r.name), axis=1)

        # 1. PAINEL GLOBAL (Para dar Baixas)
        df_painel = df_painel_raw.copy()
        df_painel['nf_ref_limpa'] = df_painel['N° da Nota fiscal'].apply(extrair_nf_painel)
        df_painel['Fornecedor_UP'] = df_painel['Fornecedor'].astype(str).str.strip().str.upper()
        
        painel_com_cnpj = pd.merge(df_painel, df_forn[['Credor_UP', 'CNPJCPF']], left_on='Fornecedor_UP', right_on='Credor_UP', how='left')
        painel_com_cnpj['CNPJCPF'] = painel_com_cnpj['CNPJCPF'].apply(limpar_cnpj)
        painel_com_cnpj['chave_p'] = painel_com_cnpj['CNPJCPF'] + "_" + painel_com_cnpj['nf_ref_limpa']
        
        chaves_lancadas_painel = set(painel_com_cnpj[painel_com_cnpj['nf_ref_limpa'] != ""]['chave_p'].unique())

        # 2. FINANCEIRO / TÍTULOS GLOBAL (Para dar Baixas por NF Lançada)
        chaves_lancadas_titulos = set()
        pedidos_com_nf_financeiro = set()
        if not df_titulos_raw.empty:
            if 'CT/OC' in df_titulos_raw.columns:
                pedidos_com_nf_financeiro = set(df_titulos_raw['CT/OC'].dropna().astype(str).str.split('.').str[0].str.strip().unique())
            
            col_nf_tit = [c for c in df_titulos_raw.columns if 'Nota' in c or 'NF' in c or 'Documento' in c or 'Nº Doc' in c]
            col_forn_tit = [c for c in df_titulos_raw.columns if 'Fornecedor' in c or 'Credor' in c or 'Razão' in c]
            
            if col_nf_tit and col_forn_tit:
                c_nf = col_nf_tit[0]
                c_forn = col_forn_tit[0]
                
                df_titulos_raw['nf_tit_limpa'] = df_titulos_raw[c_nf].apply(extrair_nf_painel)
                df_titulos_raw['Fornecedor_UP'] = df_titulos_raw[c_forn].astype(str).str.strip().str.upper()
                
                titulos_com_cnpj = pd.merge(df_titulos_raw, df_forn[['Credor_UP', 'CNPJCPF']], left_on='Fornecedor_UP', right_on='Credor_UP', how='left')
                titulos_com_cnpj['CNPJCPF'] = titulos_com_cnpj['CNPJCPF'].apply(limpar_cnpj)
                titulos_com_cnpj['chave_t'] = titulos_com_cnpj['CNPJCPF'] + "_" + titulos_com_cnpj['nf_tit_limpa']
                
                chaves_lancadas_titulos = set(titulos_com_cnpj[titulos_com_cnpj['nf_tit_limpa'] != ""]['chave_t'].unique())

        chaves_lancadas_global = chaves_lancadas_painel.union(chaves_lancadas_titulos)

        # 3. MUNDO LOCAL (PEDIDOS DA SUA OBRA)
        df_relacao = df_relacao_raw.copy()
        col_obra_rel = 'Cód. obra' if 'Cód. obra' in df_relacao.columns else df_relacao.columns[0]
        df_relacao['Cód. obra_clean'] = df_relacao[col_obra_rel].apply(limpar_cod)
        df_relacao_obra = df_relacao[df_relacao['Cód. obra_clean'] == cod_obra_alvo].copy()

        df_relacao_obra['Cód. fornecedor'] = df_relacao_obra['Cód. fornecedor'].apply(limpar_cod)
        rel_obra_com_cnpj = pd.merge(df_relacao_obra, df_forn[['Cód. Fornecedor', 'CNPJCPF']], left_on='Cód. fornecedor', right_on='Cód. Fornecedor', how='left')
        rel_obra_com_cnpj['CNPJCPF'] = rel_obra_com_cnpj['CNPJCPF'].apply(limpar_cnpj)

        cnpjs_com_pedido_na_obra = set(rel_obra_com_cnpj['CNPJCPF'].dropna().unique())

        # ABA 1: PAINEL
        resumo_painel = pd.merge(df_nf, painel_com_cnpj[['chave_p', 'N° da Nota fiscal']].drop_duplicates('chave_p'), left_on='chave_unica', right_on='chave_p', how='left')
        
        def status_painel(r):
            if r['chave_unica'] in chaves_lancadas_global: return "✅ NF Lançada"
            if r['CNPJ emitente'] in cnpjs_com_pedido_na_obra: return "⚠️ Para Verificação"
            return "❌ Sem Histórico"
        
        resumo_painel['Status'] = resumo_painel.apply(status_painel, axis=1)

        # ABA 2: PEDIDOS
        peds_agrupados = rel_obra_com_cnpj.dropna(subset=['Nº do pedido']).groupby('CNPJCPF')['Nº do pedido'].apply(lambda x: ", ".join(sorted(set(x.astype(str).unique())))).reset_index()
        resumo_pedidos = pd.merge(resumo_painel, peds_agrupados, left_on='CNPJ emitente', right_on='CNPJCPF', how='left')

        def status_pedidos(r):
            if r['Status'] == "✅ NF Lançada": return "✅ Resolvido Painel"
            if r['CNPJ emitente'] in cnpjs_com_pedido_na_obra: return "⚠️ Para Verificação"
            return "❌ Sem Histórico"
        
        resumo_pedidos['Status_Ped'] = resumo_pedidos.apply(status_pedidos, axis=1)

        # ABA 3: CONTRATO
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

        def filtrar_pedidos_por_titulo(r):
            if r['Status_CT'] == "✅ Resolvido Painel":
                return r['Nº do pedido']
                
            pedidos_string = str(r['Nº do pedido']).strip()
            if pd.isna(r['Nº do pedido']) or pedidos_string == "" or pedidos_string.lower() == "nan":
                return "NECESSÁRIO CONFECCIONAR OC"
            
            lista_pedidos = [p.split('.')[0].strip() for p in pedidos_string.split(',')]
            pedidos_em_aberto = [p for p in lista_pedidos if p not in pedidos_com_nf_financeiro]
            
            if len(pedidos_em_aberto) > 0:
                return ", ".join(pedidos_em_aberto)
            else:
                return "NECESSÁRIO CONFECCIONAR OC"

        resumo_contratos['Nº do pedido'] = resumo_contratos.apply(filtrar_pedidos_por_titulo, axis=1)

        # 🔥 --- ABA 4: EM ABERTO --- 🔥
        peds_em_aberto_lista = []

        # 1. Filtro do arquivo PAINEL (Apenas a obra do usuário)
        col_obra_painel = 'Cód. obra' if 'Cód. obra' in df_painel_raw.columns else df_painel_raw.columns[0]
        df_painel_obra = df_painel_raw[df_painel_raw[col_obra_painel].apply(limpar_cod) == cod_obra_alvo].copy()

        for _, row in df_painel_obra.iterrows():
            col_q_ped = row.iloc[16] if len(row) > 16 else None
            col_r_dt = row.iloc[17] if len(row) > 17 else None
            col_u_forn = row.iloc[20] if len(row) > 20 else ""
            col_ab_dt_nf = row.iloc[27] if len(row) > 27 else None
            col_ac_num_nf = row.iloc[28] if len(row) > 28 else None

            ped_num = str(col_q_ped).split('.')[0].strip() if pd.notna(col_q_ped) else ""
            forn_nome = str(col_u_forn).strip() if pd.notna(col_u_forn) else ""

            if ped_num != "" and ped_num.lower() != "nan" and pd.isna(col_ab_dt_nf) and pd.isna(col_ac_num_nf):
                # Forçado dayfirst=True para tratar data em formato BR (DD/MM/AAAA)
                dt_conf = pd.to_datetime(col_r_dt, dayfirst=True, errors='coerce', format='mixed')
                peds_em_aberto_lista.append({
                    'Pedido': ped_num,
                    'Data Confecção': dt_conf,
                    'Fornecedor': forn_nome
                })

        # 2. Filtro do arquivo PEDIDOS (Apenas a obra do usuário)
        for _, row in df_relacao_obra.iterrows():
            col_a_ped = row.iloc[0] if len(row) > 0 else None
            col_b_dt = row.iloc[1] if len(row) > 1 else None
            col_h_forn = row.iloc[7] if len(row) > 7 else ""
            col_as_dt_ent = row.iloc[44] if len(row) > 44 else None

            ped_num = str(col_a_ped).split('.')[0].strip() if pd.notna(col_a_ped) else ""
            forn_nome = str(col_h_forn).strip() if pd.notna(col_h_forn) else ""

            if ped_num != "" and ped_num.lower() != "nan" and pd.isna(col_as_dt_ent):
                # Forçado dayfirst=True para tratar data em formato BR (DD/MM/AAAA)
                dt_conf = pd.to_datetime(col_b_dt, dayfirst=True, errors='coerce', format='mixed')
                peds_em_aberto_lista.append({
                    'Pedido': ped_num,
                    'Data Confecção': dt_conf,
                    'Fornecedor': forn_nome
                })

        # Consolidação e Formatação da Aba 4
        if peds_em_aberto_lista:
            df_aberto = pd.DataFrame(peds_em_aberto_lista).dropna(subset=['Pedido'])
            
            df_aberto = df_aberto.groupby('Pedido').agg({
                'Data Confecção': 'min',
                'Fornecedor': lambda x: next((s for s in x if str(s).strip() != "" and str(s).lower() != "nan"), "")
            }).reset_index()
            
            df_aberto['Dias em aberto'] = (hoje - df_aberto['Data Confecção']).dt.days
            df_aberto['Data Confecção'] = df_aberto['Data Confecção'].dt.strftime('%d/%m/%Y')
            df_aberto['Dias em aberto'] = df_aberto['Dias em aberto'].fillna(0).astype(int)
            
            aba4_final = df_aberto[['Pedido', 'Data Confecção', 'Fornecedor', 'Dias em aberto']].sort_values(by='Dias em aberto', ascending=False)
        else:
            aba4_final = pd.DataFrame(columns=['Pedido', 'Data Confecção', 'Fornecedor', 'Dias em aberto'])

        # SAÍDA DAS ABAS
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
            aba4_final.to_excel(writer, sheet_name='4. EM ABERTO', index=False)
        
        st.success(f"Tudo pronto! Auditoria gerada com as datas corrigidas para o formato brasileiro na Aba 4 (Obra {cod_obra_alvo}).")
        st.download_button("📥 Baixar Auditoria", output.getvalue(), "AUDITORIA_NF_PRODUTO.xlsx")
