import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Auditoria Interna NF - Serviço", layout="wide")

st.title("📊 Auditoria Interna NF - Serviço")
st.markdown("""
### Regra de Baixas Globais:
- **Baixa Automática:** Cruzamento de NFs lançadas no **Painel** + **Relatório de Títulos/Financeiro** (Todas as Obras).
- **Status Pendentes e Pedidos:** Filtrado **exclusivamente pela SUA OBRA**.
""")

codigo_obra_usuario = st.text_input("📍 Informe o código numérico da sua obra (Ex: 2):", value="").strip()

col1, col2 = st.columns(2)
with col1:
    file_nf = st.file_uploader("1. Relatório de NF's", type=['xlsx', 'csv'])
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

def transformar_credor_limpo(df_bruto):
    if "Cód. Fornecedor" in df_bruto.columns and "Credor" in df_bruto.columns:
        return df_bruto
    for i in range(min(10, len(df_bruto))):
        row_values = [str(x).strip() for x in df_bruto.iloc[i].values]
        if 'Credor' in row_values and 'CNPJ/CPF' in row_values:
            df_header = df_bruto.iloc[i+1:].copy()
            df_header.columns = [str(c).strip() for c in df_bruto.iloc[i].values]
            df_header = df_header.loc[:, df_header.columns.notna() & (df_header.columns != 'nan')]
            def split_safe(val):
                s = str(val).strip()
                if s == "" or s == "nan": return "", ""
                if " - " in s:
                    parts = s.split(" - ")
                    return parts[0].strip(), " - ".join(parts[1:]).strip()
                return "", s
            res_split = df_header['Credor'].apply(split_safe)
            df_header['Cód. Fornecedor'] = res_split.apply(lambda x: x[0])
            df_header['Fornecedor'] = res_split.apply(lambda x: x[1])
            df_header = df_header.rename(columns={'CNPJ/CPF': 'CNPJCPF'})
            return df_header.dropna(subset=['Credor'])
    return df_bruto

if st.button("🚀 Processar Auditoria"):
    if not codigo_obra_usuario:
        st.error("⚠️ Por favor, digite o código numérico da sua obra antes de iniciar.")
        st.stop()

    if all([file_nf, file_forn, file_painel, file_relacao, file_contrato, file_titulo]):
        cod_obra_alvo = limpar_cod(codigo_obra_usuario)

        df_nf = pd.read_excel(file_nf) if file_nf.name.endswith('xlsx') else pd.read_csv(file_nf)
        df_forn_raw = pd.read_excel(file_forn, header=None) if file_forn.name.endswith('xlsx') else pd.read_csv(file_forn, header=None)
        df_painel = pd.read_excel(file_painel) if file_painel.name.endswith('xlsx') else pd.read_csv(file_painel)
        df_relacao = pd.read_excel(file_relacao) if file_relacao.name.endswith('xlsx') else pd.read_csv(file_relacao)
        df_bruto_ct = pd.read_excel(file_contrato, header=None) if file_contrato.name.endswith('xlsx') else pd.read_csv(file_contrato, header=None)
        df_titulos_raw = estruturar_titulo_limpo(file_titulo)

        df_forn = transformar_credor_limpo(df_forn_raw)
        df_forn['CNPJCPF'] = df_forn['CNPJCPF'].astype(str).apply(limpar_cnpj)
        df_forn['Credor_UP'] = df_forn['Credor'].astype(str).str.strip().str.upper()

        NF_CNPJ = 'CNPJ emitente' if 'CNPJ emitente' in df_nf.columns else ('CNPJ Prestador (CNPJ)' if 'CNPJ Prestador (CNPJ)' in df_nf.columns else df_nf.columns[1])
        NF_NUMERO = 'Número/Série' if 'Número/Série' in df_nf.columns else ('Número NFS-e (nNFSe)' if 'Número NFS-e (nNFSe)' in df_nf.columns else df_nf.columns[11])
        NF_FORN = 'Emitente' if 'Emitente' in df_nf.columns else ('Nome Prestador (xNome)' if 'Nome Prestador (xNome)' in df_nf.columns else df_nf.columns[0])
        NF_DATA = 'Emissão' if 'Emissão' in df_nf.columns else ('Data/Hora Emissão DPS (dhEmi)' if 'Data/Hora Emissão DPS (dhEmi)' in df_nf.columns else df_nf.columns[9])
        NF_VALOR = 'Valor' if 'Valor' in df_nf.columns else ('Valor do Serviço (vServ) (vServ)' if 'Valor do Serviço (vServ) (vServ)' in df_nf.columns else df_nf.columns[12])
        
        df_nf[NF_CNPJ] = df_nf[NF_CNPJ].astype(str).apply(limpar_cnpj)
        
        def extrair_numero_nf_puro(x):
            texto = str(x).strip()
            if not texto or texto == "nan": return ""
            parte_antes_da_barra = texto.split('/')[0]
            return "".join(filter(str.isdigit, parte_antes_da_barra)).strip()

        df_nf['nf_limpa'] = df_nf[NF_NUMERO].apply(extrair_numero_nf_puro)
        df_nf['chave_unica'] = df_nf[NF_CNPJ] + "_" + df_nf['nf_limpa']

        # 1. PAINEL GLOBAL (Para dar baixas)
        df_painel['Fornecedor_UP'] = df_painel['Fornecedor'].astype(str).str.strip().str.upper()
        painel_com_cnpj = pd.merge(df_painel, df_forn[['Credor_UP', 'CNPJCPF']], left_on='Fornecedor_UP', right_on='Credor_UP', how='left')
        painel_com_cnpj['CNPJCPF'] = painel_com_cnpj['CNPJCPF'].apply(limpar_cnpj)
        
        col_nf_painel = 'N° da Nota fiscal' if 'N° da Nota fiscal' in df_painel.columns else ('N° da Nota Fiscal' if 'N° da Nota Fiscal' in df_painel.columns else 'N. do Pedido')
        
        def extrair_numero_nf_painel(x):
            texto = str(x).strip()
            if not texto or texto == "nan": return ""
            parte_final = texto.split('/')[-1]
            return "".join(filter(str.isdigit, parte_final)).strip()

        painel_com_cnpj['nf_ref_limpa'] = painel_com_cnpj[col_nf_painel].apply(extrair_numero_nf_painel)
        painel_com_cnpj['chave_p'] = painel_com_cnpj['CNPJCPF'] + "_" + painel_com_cnpj['nf_ref_limpa']
        
        chaves_lancadas_painel = set(painel_com_cnpj[painel_com_cnpj['nf_ref_limpa'] != ""]['chave_p'].unique())

        # 2. FINANCEIRO / TÍTULOS GLOBAL (Para dar baixas por NF Lançada)
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
                
                df_titulos_raw['nf_tit_limpa'] = df_titulos_raw[c_nf].apply(extrair_numero_nf_painel)
                df_titulos_raw['Fornecedor_UP'] = df_titulos_raw[c_forn].astype(str).str.strip().str.upper()
                
                titulos_com_cnpj = pd.merge(df_titulos_raw, df_forn[['Credor_UP', 'CNPJCPF']], left_on='Fornecedor_UP', right_on='Credor_UP', how='left')
                titulos_com_cnpj['CNPJCPF'] = titulos_com_cnpj['CNPJCPF'].apply(limpar_cnpj)
                titulos_com_cnpj['chave_t'] = titulos_com_cnpj['CNPJCPF'] + "_" + titulos_com_cnpj['nf_tit_limpa']
                
                chaves_lancadas_titulos = set(titulos_com_cnpj[titulos_com_cnpj['nf_tit_limpa'] != ""]['chave_t'].unique())

        # UNIFICAÇÃO DAS BAIXAS (Painel + Títulos)
        chaves_lancadas_global = chaves_lancadas_painel.union(chaves_lancadas_titulos)

        # 3. MUNDO LOCAL (PEDIDOS DA SUA OBRA)
        col_obra_rel = 'Cód. obra' if 'Cód. obra' in df_relacao.columns else df_relacao.columns[0]
        df_relacao['Cód. obra_clean'] = df_relacao[col_obra_rel].apply(limpar_cod)
        df_relacao_obra = df_relacao[df_relacao['Cód. obra_clean'] == cod_obra_alvo].copy()

        df_relacao_obra['Cód. fornecedor'] = df_relacao_obra['Cód. fornecedor'].apply(limpar_cod)
        rel_obra_com_cnpj = pd.merge(df_relacao_obra, df_forn[['Cód. Fornecedor', 'CNPJCPF']], left_on='Cód. fornecedor', right_on='Cód. Fornecedor', how='left')
        rel_obra_com_cnpj['CNPJCPF'] = rel_obra_com_cnpj['CNPJCPF'].apply(limpar_cnpj)

        cnpjs_com_pedido_na_obra = set(rel_obra_com_cnpj['CNPJCPF'].dropna().unique())

        # ABA 1: PAINEL
        resumo_painel = pd.merge(df_nf, painel_com_cnpj[['chave_p', col_nf_painel]].drop_duplicates('chave_p'), left_on='chave_unica', right_on='chave_p', how='left')
        if col_nf_painel in resumo_painel.columns and col_nf_painel != 'N° da Nota fiscal':
            resumo_painel = resumo_painel.rename(columns={col_nf_painel: 'N° da Nota fiscal'})
        elif 'N° da Nota fiscal' not in resumo_painel.columns:
            resumo_painel['N° da Nota fiscal'] = ""

        def definir_status_painel(r):
            if r['chave_unica'] in chaves_lancadas_global: return "✅ NF Lançada"
            if r[NF_CNPJ] in cnpjs_com_pedido_na_obra: return "⚠️ Para Verificação"
            return "❌ Sem Histórico"
        resumo_painel['Status'] = resumo_painel.apply(definir_status_painel, axis=1)
        
        # ABA 2: PEDIDOS
        peds_agrupados = rel_obra_com_cnpj.dropna(subset=['Nº do pedido']).groupby('CNPJCPF')['Nº do pedido'].apply(lambda x: ", ".join(sorted(set(x.astype(str).unique())))).reset_index()
        resumo_pedidos = pd.merge(resumo_painel, peds_agrupados, left_on=NF_CNPJ, right_on='CNPJCPF', how='left')

        def status_pedidos(r):
            if r['chave_unica'] in chaves_lancadas_global: return "✅ Resolvido Painel"
            if r[NF_CNPJ] in cnpjs_com_pedido_na_obra: return "⚠️ Para Verificação"
            return "❌ Sem Histórico"
        resumo_pedidos['Status_Ped'] = resumo_pedidos.apply(status_pedidos, axis=1)

        # ABA 3: CONTRATO
        registros_ct = []
        item_atual = {'Contrato': None, 'CNPJ': None}
        for i in range(len(df_bruto_ct)):
            linha = df_bruto_ct.iloc[i]
            col_a = str(linha[0]).strip() if pd.notna(linha[0]) else ""
            col_d = linha[3] if pd.notna(linha[3]) else "" 
            if col_a == "Contrato":
                item_atual['Contrato'] = str(col_d).strip()
            elif col_a == "CNPJ" and item_atual['Contrato']:
                item_atual['CNPJ'] = limpar_cnpj(col_d)
                registros_ct.append(item_atual.copy())

        cts_agrupados = pd.DataFrame(registros_ct).groupby('CNPJ')['Contrato'].apply(lambda x: ", ".join(sorted(set(x.astype(str).unique())))).reset_index() if registros_ct else pd.DataFrame(columns=['CNPJ', 'Contrato'])

        resumo_contratos = pd.merge(resumo_pedidos, cts_agrupados, left_on=NF_CNPJ, right_on='CNPJ', how='left')
        
        def status_contratos(r):
            if r['chave_unica'] in chaves_lancadas_global: return "✅ Resolvido Painel"
            if pd.notna(r['Contrato']) and str(r['Contrato']).strip() != "": return "📄 Vínculo Contratual"
            return r['Status_Ped']
        resumo_contratos['Status_CT'] = resumo_contratos.apply(status_contratos, axis=1)

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

        aba3_final = resumo_contratos[[
            NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, 
            'N° da Nota fiscal', 'Nº do pedido', 'Contrato', 'Status_CT'
        ]].rename(columns={'Status_CT': 'Status', 'Nº do pedido': 'Pedido'})

        # SAÍDA
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resumo_painel[[NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, 'N° da Nota fiscal', 'Status']].to_excel(writer, sheet_name='1. PAINEL', index=False)
            cols_aba2 = [NF_NUMERO, NF_CNPJ, NF_FORN, NF_DATA, NF_VALOR, 'N° da Nota fiscal', 'Nº do pedido', 'Status_Ped']
            resumo_pedidos[cols_aba2].rename(columns={'Status_Ped': 'Status', 'Nº do pedido': 'Pedido'}).to_excel(writer, sheet_name='2. PEDIDOS', index=False)
            aba3_final.to_excel(writer, sheet_name='3. CONTRATO', index=False)
        
        st.success(f"Tudo pronto! Auditoria gerada focada na Obra {cod_obra_alvo}.")
        st.download_button(label="📥 Baixar Auditoria", data=output.getvalue(), file_name="AUDITORIA_NF_SERVIÇO.xlsx")
