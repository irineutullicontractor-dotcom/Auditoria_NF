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
- **Aba 4. EM ABERTO:** Consolidação de pedidos ativos da sua obra que ainda não possuem Nota Fiscal/Entrada vinculada (com Fornecedor, Dias em Aberto e Status do Pedido, desconsiderando **Cancelados**).
### Instruções de uso:
1. Carregue o relatório de **NF's** - Puxar relatório do mês vigente (escolher a empresa pertinente ao CNPJ).
2. Carregue o relatório de **Credores**.
3. Carregue o relatório do **Painel** - Puxar relatório de no mínimo 90 dias atrás até a data vigente.
4. Carregue o relatório de **Pedidos** - Puxar relatório de no mínimo 90 dias atrás até a data vigente.
5. Carregue o relatório de **Contratos** - Puxar relatório de 01/01/2020 até a data vigente.
""")

codigo_obra_usuario = st.text_input("📍 Informe o código numérico da sua obra (Ex: 2):", value="").strip()

col1, col2 = st.columns(2)
with col1:
    file_nf_prod = st.file_uploader("1. Relatório de NF's - Home / Notas Fiscais / Recepção de NF-e / Relatórios / Notas Fiscais Recebidas.", type=['xlsx'])
    file_forn = st.file_uploader("2. Relatório de Credores - Home / Mais Opções / Apoio / Relatórios / Pessoas / Credores.", type=['xlsx'])
    file_painel = st.file_uploader("3. Relatório Painel - Home / Suprimentos / Compras / Painel de Compras (Novo).", type=['xlsx'])
with col2:
    file_relacao = st.file_uploader("4. Relatório Pedidos - Home / Suprimentos / Compras / Relatórios / Pedidos de compra / Relação de Pedidos de Compra (Novo).", type=['xlsx'])
    file_contrato = st.file_uploader("5. Relatório Contrato - Home / Suprimentos / Contratos e Medições / Relatórios / Contratos / Emissão de Contratos.", type=['xlsx'])
    file_titulo = st.file_uploader("6. Relatório Titulo - Home / Financeiro / Contas a Pagar / Relatórios / Títulos por Data.", type=['xlsx'])

# ==============================================================================
# --- FUNÇÕES AUXILIARES (HELPERS) ---
# ==============================================================================

def apenas_numeros(v):
    if pd.isna(v): return ""
    return "".join(filter(str.isdigit, str(v)))

def limpar_cnpj_cpf(v):
    num = apenas_numeros(v)
    if not num: return ""
    return num.zfill(14) if len(num) > 11 else num.zfill(11)

def extrair_raiz_cnpj(cnpj_clean):
    c = str(cnpj_clean).strip()
    return c[:8] if len(c) >= 8 else c

def limpar_cod(v):
    if pd.isna(v): return ""
    s = str(v).strip()
    if " - " in s:
        s = s.split(" - ")[0].strip()
    return s.split('.')[0].lstrip('0')

def extrair_numero_apos_barra(v):
    if pd.isna(v) or str(v).strip() == "" or str(v).lower() == "nan": return ""
    s = str(v).strip()
    if '/' in s:
        s = s.split('/')[-1]
    num = apenas_numeros(s)
    return num.lstrip('0')

def extrair_numero_antes_barra(v):
    if pd.isna(v) or str(v).strip() == "" or str(v).lower() == "nan": return ""
    s = str(v).strip()
    if '/' in s:
        s = s.split('/')[0]
    num = apenas_numeros(s)
    return num.lstrip('0')

def encontrar_coluna(df, termos_busca):
    for col in df.columns:
        col_str = str(col).strip().lower()
        for termo in termos_busca:
            if termo in col_str:
                return col
    return None

def campo_preenchido(val):
    if pd.isna(val): return False
    s = str(val).strip().lower()
    return s not in ["", "nan", "none", "null", "0", "00/00/0000", "00/00/00"]

# ==============================================================================
# --- LEITURA DE ARQUIVOS ---
# ==============================================================================

def carregar_df(file, skiprows=0, header=0):
    if file.name.endswith('.csv'):
        return pd.read_csv(file, skiprows=skiprows, header=header)
    return pd.read_excel(file, skiprows=skiprows, header=header)

def estruturar_credor(file):
    df_bruto = pd.read_excel(file, header=None) if file.name.endswith('.xlsx') else pd.read_csv(file, header=None)
    for i in range(min(15, len(df_bruto))):
        row_values = [str(x).strip() for x in df_bruto.iloc[i].values]
        if 'Credor' in row_values and ('CNPJ/CPF' in row_values or 'CNPJ' in row_values):
            df_header = df_bruto.iloc[i+1:].copy()
            df_header.columns = [str(c).strip() for c in df_bruto.iloc[i].values]
            
            col_cred = 'Credor'
            col_cnpj = 'CNPJ/CPF' if 'CNPJ/CPF' in df_header.columns else 'CNPJ'
            
            df_header['Cod_Forn'] = df_header[col_cred].apply(limpar_cod)
            df_header['CNPJ_Clean'] = df_header[col_cnpj].apply(limpar_cnpj_cpf)
            return df_header
    return pd.DataFrame()

def estruturar_nf_produto(file):
    df_bruto = pd.read_excel(file, header=None) if file.name.endswith('.xlsx') else pd.read_csv(file, header=None)
    registros = []
    cnpj_dest = None
    colunas_id = None
    processando = False

    for i, row in df_bruto.iterrows():
        val_a = str(row[0]).strip() if pd.notna(row[0]) else ""
        if "CNPJ do destinatário:" in val_a:
            cnpj_dest = limpar_cnpj_cpf(row[3])
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

def estruturar_titulo_limpo(file):
    df_bruto = pd.read_excel(file, header=None) if file.name.endswith('.xlsx') else pd.read_csv(file, header=None)
    inicio = None
    for i, row in df_bruto.iterrows():
        if str(row[0]).strip().lower() == "item":
            inicio = i
            break
    if inicio is None: return pd.DataFrame()
    
    df = carregar_df(file, skiprows=inicio)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# ==============================================================================
# --- EXECUÇÃO DA AUDITORIA ---
# ==============================================================================

if st.button("🚀 Iniciar Auditoria"):
    if not codigo_obra_usuario:
        st.error("⚠️ Por favor, digite o código numérico da sua obra antes de iniciar.")
        st.stop()

    if all([file_nf_prod, file_forn, file_painel, file_relacao, file_contrato, file_titulo]):
        cod_obra_alvo = limpar_cod(codigo_obra_usuario)
        hoje = pd.to_datetime(datetime.date.today())

        # 1. FORNECEDORES
        df_credor = estruturar_credor(file_forn)
        mapa_cod_to_cnpj = dict(zip(df_credor['Cod_Forn'], df_credor['CNPJ_Clean']))

        # 2. RELATÓRIO NF PRODUTO
        df_nf = estruturar_nf_produto(file_nf_prod)
        df_nf['CNPJ emitente'] = df_nf['CNPJ emitente'].apply(limpar_cnpj_cpf)
        df_nf['raiz_cnpj'] = df_nf['CNPJ emitente'].apply(extrair_raiz_cnpj)
        df_nf['nf_limpa'] = df_nf['Núm/Série'].apply(extrair_numero_antes_barra)
        df_nf['chave_unica'] = df_nf['CNPJ emitente'] + "_" + df_nf['nf_limpa']
        df_nf['chave_raiz'] = df_nf['raiz_cnpj'] + "_" + df_nf['nf_limpa']

        # 3. PAINEL DE COMPRAS
        df_painel = carregar_df(file_painel)
        df_painel.columns = [str(c).strip() for c in df_painel.columns]
        
        col_forn_p = encontrar_coluna(df_painel, ['fornecedor', 'credor', 'razão social']) or df_painel.columns[0]
        col_nf_p = encontrar_coluna(df_painel, ['n° da nota fiscal', 'nota fiscal', 'n° da nota', 'nfe']) or df_painel.columns[1]
        col_obra_p = encontrar_coluna(df_painel, ['obra', 'cód. obra', 'cod obra']) or df_painel.columns[2]
        col_ped_p = encontrar_coluna(df_painel, ['n° do pedido', 'pedido', 'nº pedido']) or df_painel.columns[3]

        df_painel['Cod_Forn'] = df_painel[col_forn_p].apply(limpar_cod)
        df_painel['CNPJ_Painel'] = df_painel['Cod_Forn'].map(mapa_cod_to_cnpj).fillna("")
        df_painel['raiz_cnpj'] = df_painel['CNPJ_Painel'].apply(extrair_raiz_cnpj)
        df_painel['nf_limpa'] = df_painel[col_nf_p].apply(extrair_numero_apos_barra)
        
        df_painel['chave_p'] = df_painel['CNPJ_Painel'] + "_" + df_painel['nf_limpa']
        df_painel['chave_raiz_p'] = df_painel['raiz_cnpj'] + "_" + df_painel['nf_limpa']
        df_painel['Cod_Obra_Clean'] = df_painel[col_obra_p].apply(limpar_cod)

        # Chaves de NFs presentes no Painel (Geral)
        painel_com_nf = df_painel[df_painel['nf_limpa'] != ""]
        chaves_lancadas_painel = set(painel_com_nf['chave_p'].unique())
        chaves_raiz_painel = set(painel_com_nf['chave_raiz_p'].unique())

        # 4. TÍTULOS FINANCEIROS
        df_titulos = estruturar_titulo_limpo(file_titulo)
        chaves_lancadas_titulos = set()
        chaves_raiz_titulos = set()
        pedidos_com_nf_financeiro = set()

        if not df_titulos.empty:
            df_titulos['Cod_Forn'] = df_titulos['Credor'].apply(limpar_cod)
            df_titulos['CNPJ_Tit'] = df_titulos['Cod_Forn'].map(mapa_cod_to_cnpj).fillna("")
            df_titulos['raiz_cnpj'] = df_titulos['CNPJ_Tit'].apply(extrair_raiz_cnpj)
            df_titulos['nf_limpa'] = df_titulos['Documento'].apply(extrair_numero_apos_barra)
            
            df_titulos['chave_t'] = df_titulos['CNPJ_Tit'] + "_" + df_titulos['nf_limpa']
            df_titulos['chave_raiz_t'] = df_titulos['raiz_cnpj'] + "_" + df_titulos['nf_limpa']
            
            titulos_com_nf = df_titulos[df_titulos['nf_limpa'] != ""]
            chaves_lancadas_titulos = set(titulos_com_nf['chave_t'].unique())
            chaves_raiz_titulos = set(titulos_com_nf['chave_raiz_t'].unique())
            
            if 'CT/OC' in df_titulos.columns:
                pedidos_com_nf_financeiro = set(df_titulos['CT/OC'].dropna().astype(str).apply(limpar_cod).unique())

        # CONSOLIDAÇÃO GLOBAL DE NF LANÇADAS (PAINEL + FINANCEIRO)
        chaves_lancadas_global = chaves_lancadas_painel.union(chaves_lancadas_titulos)
        chaves_raiz_global = chaves_raiz_painel.union(chaves_raiz_titulos)

        # 5. HISTÓRICO DA OBRA
        df_relacao = carregar_df(file_relacao)
        df_relacao.columns = [str(c).strip() for c in df_relacao.columns]

        col_obra_rel = encontrar_coluna(df_relacao, ['cód. obra', 'obra', 'cod obra']) or df_relacao.columns[0]
        col_ped_rel = encontrar_coluna(df_relacao, ['n° do pedido', 'pedido', 'nº pedido']) or 'Pedido'
        col_forn_rel = encontrar_coluna(df_relacao, ['cód. fornecedor', 'fornecedor', 'cod. fornecedor']) or df_relacao.columns[1]

        df_relacao['Obra_Clean'] = df_relacao[col_obra_rel].apply(limpar_cod)
        df_relacao_obra = df_relacao[df_relacao['Obra_Clean'] == cod_obra_alvo].copy()

        df_relacao_obra['Cod_Forn'] = df_relacao_obra[col_forn_rel].apply(limpar_cod)
        df_relacao_obra['CNPJ_Forn'] = df_relacao_obra['Cod_Forn'].map(mapa_cod_to_cnpj).fillna("")

        cnpjs_obra_pedidos = set(df_relacao_obra['CNPJ_Forn'].dropna().unique())
        cnpjs_obra_painel = set(df_painel[df_painel['Cod_Obra_Clean'] == cod_obra_alvo]['CNPJ_Painel'].dropna().unique())
        cnpjs_com_historico_na_obra = cnpjs_obra_pedidos.union(cnpjs_obra_painel)

        # --- ABA 1: PAINEL ---
        resumo_painel = pd.merge(
            df_nf, 
            df_painel[['chave_p', col_nf_p]].drop_duplicates('chave_p'), 
            left_on='chave_unica', 
            right_on='chave_p', 
            how='left'
        )
        if col_nf_p != 'N° da Nota fiscal':
            resumo_painel.rename(columns={col_nf_p: 'N° da Nota fiscal'}, inplace=True)

        def definir_status_painel(r):
            # 1. Se foi lançada (via Painel ou Financeiro)
            if r['chave_unica'] in chaves_lancadas_global:
                return "✅ NF Lançada"
            # 2. Se foi lançada sob outra filial (mesma raiz CNPJ)
            if r['chave_raiz'] in chaves_raiz_global:
                return "🚨 DIVERGÊNCIA MATRIZ/FILIAL (NF em outro CNPJ)"
            # 3. Fornecedor possui histórico na obra
            if r['CNPJ emitente'] in cnpjs_com_historico_na_obra:
                return "⚠️ Para Verificação"
            return "❌ Sem Histórico"

        resumo_painel['Status'] = resumo_painel.apply(definir_status_painel, axis=1)

        # --- ABA 2: PEDIDOS ---
        peds_painel = df_painel[df_painel['Cod_Obra_Clean'] == cod_obra_alvo].groupby('CNPJ_Painel')[col_ped_p].apply(lambda x: list(x.dropna().astype(str))).to_dict()
        peds_rel = df_relacao_obra.groupby('CNPJ_Forn')[col_ped_rel].apply(lambda x: list(x.dropna().astype(str))).to_dict() if col_ped_rel in df_relacao_obra.columns else {}

        def consolidar_pedidos_obra(cnpj):
            l1 = peds_painel.get(cnpj, [])
            l2 = peds_rel.get(cnpj, [])
            todos = sorted(set([limpar_cod(p) for p in l1 + l2 if limpar_cod(p) != ""]))
            return ", ".join(todos) if todos else ""

        resumo_pedidos = resumo_painel.copy()
        resumo_pedidos['Pedido'] = resumo_pedidos['CNPJ emitente'].apply(consolidar_pedidos_obra)
        resumo_pedidos['Status_Ped'] = resumo_pedidos.apply(lambda r: "✅ Resolvido Painel" if r['Status'] == "✅ NF Lançada" else r['Status'], axis=1)

        # --- ABA 3: CONTRATO ---
        df_contrato_bruto = carregar_df(file_contrato, header=None)
        registros_ct = []
        item_ct = {'Contrato': None, 'CNPJ': None}
        for i in range(len(df_contrato_bruto)):
            l = df_contrato_bruto.iloc[i]
            ca = str(l[0]).strip() if pd.notna(l[0]) else ""
            if ca == "Contrato": item_ct['Contrato'] = str(l[3]).strip()
            elif ca == "CNPJ" and item_ct['Contrato']:
                item_ct['CNPJ'] = limpar_cnpj_cpf(l[3])
                registros_ct.append(item_ct.copy())

        cts_agrupados = pd.DataFrame(registros_ct).groupby('CNPJ')['Contrato'].apply(lambda x: ", ".join(sorted(set(x.astype(str).unique())))).reset_index() if registros_ct else pd.DataFrame(columns=['CNPJ', 'Contrato'])

        resumo_contratos = pd.merge(resumo_pedidos, cts_agrupados, left_on='CNPJ emitente', right_on='CNPJ', how='left')
        titulos_map = df_titulos[['chave_t', 'Documento']].drop_duplicates('chave_t').set_index('chave_t')['Documento'].to_dict() if not df_titulos.empty else {}

        def definir_status_contrato(r):
            if r['chave_unica'] in chaves_lancadas_titulos or r['Status_Ped'] == "✅ Resolvido Painel":
                return "✅ NF Lançada"
            if r['Status_Ped'] == "🚨 DIVERGÊNCIA MATRIZ/FILIAL (NF em outro CNPJ)":
                return "🚨 DIVERGÊNCIA MATRIZ/FILIAL (NF em outro CNPJ)"
            if pd.notna(r['Contrato']) and str(r['Contrato']).strip() != "":
                return "📄 Vínculo Contratual"
            return r['Status_Ped']

        resumo_contratos['Status_CT'] = resumo_contratos.apply(definir_status_contrato, axis=1)
        resumo_contratos['N° da Nota fiscal'] = resumo_contratos.apply(
            lambda r: titulos_map.get(r['chave_unica'], r['N° da Nota fiscal']), axis=1
        )

        def filtrar_pedidos_fin(r):
            if r['Status_CT'] in ["✅ Resolvido Painel", "✅ NF Lançada"]:
                return r['Pedido']
            peds = [p.strip() for p in str(r['Pedido']).split(',') if p.strip()]
            em_aberto = [p for p in peds if p not in pedidos_com_nf_financeiro]
            return ", ".join(em_aberto) if em_aberto else "NECESSÁRIO CONFECCIONAR OC"

        resumo_contratos['Pedido'] = resumo_contratos.apply(filtrar_pedidos_fin, axis=1)

        # --- ABA 4: EM ABERTO ---
        peds_aberto = []
        status_validos = ['pendente', 'parcialmente']

        col_st_p = encontrar_coluna(df_painel, ['situação do pedido', 'situação', 'status'])
        col_dt_p = encontrar_coluna(df_painel, ['data do pedido', 'data pedido', 'emissão'])
        col_nf_p_painel = encontrar_coluna(df_painel, ['n° da nota fiscal', 'nota fiscal', 'n° nf', 'nfe'])
        col_forn_p_painel = encontrar_coluna(df_painel, ['fornecedor', 'credor', 'razão social'])

        for _, row in df_painel[df_painel['Cod_Obra_Clean'] == cod_obra_alvo].iterrows():
            st_p = str(row.get(col_st_p, '')).strip().lower() if col_st_p else ""
            if any(sv in st_p for sv in status_validos) and "totalmente" not in st_p:
                num_nf = row.get(col_nf_p_painel, None) if col_nf_p_painel else None
                if not campo_preenchido(num_nf):
                    ped_n = limpar_cod(row.get(col_ped_p, ''))
                    if ped_n:
                        dt_c = pd.to_datetime(row.get(col_dt_p, None), dayfirst=True, errors='coerce') if col_dt_p else None
                        forn_nome = row.get(col_forn_p_painel, '') if col_forn_p_painel else ''
                        peds_aberto.append({
                            'Pedido': ped_n,
                            'Data Confecção': dt_c,
                            'Fornecedor': forn_nome,
                            'Status Pedido': str(row.get(col_st_p, '')).strip().capitalize()
                        })

        col_st_rel = encontrar_coluna(df_relacao_obra, ['status entrega', 'status', 'situação'])
        col_dt_rel = encontrar_coluna(df_relacao_obra, ['data pedido', 'data emissao', 'emissão'])
        col_dt_ent_rel = encontrar_coluna(df_relacao_obra, ['data entregue', 'data entrega', 'entregue'])
        col_forn_rel_nome = encontrar_coluna(df_relacao_obra, ['fornecedor', 'credor', 'razão social'])

        for _, row in df_relacao_obra.iterrows():
            st_rel = str(row.get(col_st_rel, '')).strip().lower() if col_st_rel else ""
            if any(sv in st_rel for sv in status_validos) and "totalmente" not in st_rel:
                dt_ent = row.get(col_dt_ent_rel, None) if col_dt_ent_rel else None
                if not campo_preenchido(dt_ent):
                    ped_n = limpar_cod(row.get(col_ped_rel, ''))
                    if ped_n:
                        dt_c = pd.to_datetime(row.get(col_dt_rel, None), dayfirst=True, errors='coerce') if col_dt_rel else None
                        forn_nome = row.get(col_forn_rel_nome, row.get(col_forn_rel, '')) if col_forn_rel_nome else ''
                        peds_aberto.append({
                            'Pedido': ped_n,
                            'Data Confecção': dt_c,
                            'Fornecedor': forn_nome,
                            'Status Pedido': str(row.get(col_st_rel, '')).strip().capitalize()
                        })

        if peds_aberto:
            df_aberto = pd.DataFrame(peds_aberto).dropna(subset=['Pedido'])
            df_aberto = df_aberto.groupby('Pedido').agg({
                'Data Confecção': 'min',
                'Fornecedor': 'first',
                'Status Pedido': 'first'
            }).reset_index()

            df_aberto['Dias em aberto'] = (hoje - df_aberto['Data Confecção']).dt.days.fillna(0).astype(int)
            df_aberto['Data Confecção'] = df_aberto['Data Confecção'].dt.strftime('%d/%m/%Y')
            aba4_final = df_aberto[['Pedido', 'Data Confecção', 'Fornecedor', 'Dias em aberto', 'Status Pedido']].sort_values(by='Dias em aberto', ascending=False)
        else:
            aba4_final = pd.DataFrame(columns=['Pedido', 'Data Confecção', 'Fornecedor', 'Dias em aberto', 'Status Pedido'])

        # EXPORTAÇÃO
        cols_base = ['Núm/Série', 'CNPJ emitente', 'Emitente', 'Emissão', 'Valor']
        cols_extra = ['CNPJ Destinatário', 'Destinatário']

        aba1 = resumo_painel[cols_base + ['N° da Nota fiscal', 'Status'] + cols_extra]
        aba2 = resumo_pedidos[cols_base + ['N° da Nota fiscal', 'Pedido', 'Status_Ped'] + cols_extra].rename(columns={'Status_Ped': 'Status'})
        aba3 = resumo_contratos[cols_base + ['N° da Nota fiscal', 'Pedido', 'Contrato', 'Status_CT'] + cols_extra].rename(columns={'Status_CT': 'Status'})

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            aba1.to_excel(writer, sheet_name='1. PAINEL', index=False)
            aba2.to_excel(writer, sheet_name='2. PEDIDOS', index=False)
            aba3.to_excel(writer, sheet_name='3. CONTRATO', index=False)
            aba4_final.to_excel(writer, sheet_name='4. EM ABERTO', index=False)

        st.success(f" Auditoria gerada com sucesso para a Obra {cod_obra_alvo}!")
        st.download_button("📥 Baixar Relatório Auditoria", output.getvalue(), "AUDITORIA_NF_PRODUTO.xlsx")
