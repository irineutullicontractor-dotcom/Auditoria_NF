import streamlit as st
import pandas as pd
import datetime
import io
import re

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

# --- FUNÇÕES DE SANITIZAÇÃO E LIMPEZA ---

def apenas_numeros(v):
    if pd.isna(v): return ""
    return "".join(filter(str.isdigit, str(v)))

def limpar_cnpj_cpf(v):
    num = apenas_numeros(v)
    if not num: return ""
    return num.zfill(14) if len(num) > 11 else num.zfill(11)

def limpar_cod(v):
    if pd.isna(v): return ""
    s = str(v).strip()
    if " - " in s:
        s = s.split(" - ")[0].strip()
    return s.split('.')[0].lstrip('0')

def extrair_numero_apos_barra(v):
    """Usado para Painel (NFE/280680) e Títulos (NF/38, NF / 38)"""
    if pd.isna(v) or str(v).strip() == "" or str(v).lower() == "nan": return ""
    s = str(v).strip()
    if '/' in s:
        s = s.split('/')[-1]
    num = apenas_numeros(s)
    return num.lstrip('0')

def extrair_numero_antes_barra(v):
    """Usado para NF Produto (Núm/Série: 34994/1)"""
    if pd.isna(v) or str(v).strip() == "" or str(v).lower() == "nan": return ""
    s = str(v).strip()
    if '/' in s:
        s = s.split('/')[0]
    num = apenas_numeros(s)
    return num.lstrip('0')

# --- CARREGADORES DE ARQUIVOS ROBUSTOS ---

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

# --- PROCESSAMENTO DA AUDITORIA ---

if st.button("🚀 Iniciar Auditoria"):
    if not codigo_obra_usuario:
        st.error("⚠️ Por favor, digite o código numérico da sua obra antes de iniciar.")
        st.stop()

    if all([file_nf_prod, file_forn, file_painel, file_relacao, file_contrato, file_titulo]):
        cod_obra_alvo = limpar_cod(codigo_obra_usuario)
        hoje = pd.to_datetime(datetime.date.today())

        # 1. DE-PARA GLOBAL DE FORNECEDORES (CÓDIGO <-> CNPJ)
        df_credor = estruturar_credor(file_forn)
        mapa_cod_to_cnpj = dict(zip(df_credor['Cod_Forn'], df_credor['CNPJ_Clean']))
        mapa_cnpj_to_cod = dict(zip(df_credor['CNPJ_Clean'], df_credor['Cod_Forn']))

        # 2. CARREGAR E TRATAR NF PRODUTO
        df_nf = estruturar_nf_produto(file_nf_prod)
        df_nf['CNPJ emitente'] = df_nf['CNPJ emitente'].apply(limpar_cnpj_cpf)
        df_nf['nf_limpa'] = df_nf['Núm/Série'].apply(extrair_numero_antes_barra)
        df_nf['chave_unica'] = df_nf['CNPJ emitente'] + "_" + df_nf['nf_limpa']

        # 3. TRATAR PAINEL DE COMPRAS
        df_painel = carregar_df(file_painel)
        df_painel['Cod_Forn'] = df_painel['Fornecedor'].apply(limpar_cod)
        df_painel['CNPJ_Painel'] = df_painel['Cod_Forn'].map(mapa_cod_to_cnpj).fillna("")
        df_painel['nf_limpa'] = df_painel['N° da Nota fiscal'].apply(extrair_numero_apos_barra)
        df_painel['chave_p'] = df_painel['CNPJ_Painel'] + "_" + df_painel['nf_limpa']
        df_painel['Cod_Obra_Clean'] = df_painel['Obra'].apply(limpar_cod)

        # Chaves de NFs Lançadas no Painel (Global)
        chaves_lancadas_painel = set(df_painel[df_painel['nf_limpa'] != ""]['chave_p'].unique())

        # 4. TRATAR TÍTULOS FINANCEIROS
        df_titulos = estruturar_titulo_limpo(file_titulo)
        chaves_lancadas_titulos = set()
        pedidos_com_nf_financeiro = set()

        if not df_titulos.empty:
            df_titulos['Cod_Forn'] = df_titulos['Credor'].apply(limpar_cod)
            df_titulos['CNPJ_Tit'] = df_titulos['Cod_Forn'].map(mapa_cod_to_cnpj).fillna("")
            df_titulos['nf_limpa'] = df_titulos['Documento'].apply(extrair_numero_apos_barra)
            df_titulos['chave_t'] = df_titulos['CNPJ_Tit'] + "_" + df_titulos['nf_limpa']
            
            chaves_lancadas_titulos = set(df_titulos[df_titulos['nf_limpa'] != ""]['chave_t'].unique())
            if 'CT/OC' in df_titulos.columns:
                pedidos_com_nf_financeiro = set(df_titulos['CT/OC'].dropna().astype(str).apply(limpar_cod).unique())

        # CHAVES GLOBAIS DE BAIXA
        chaves_lancadas_global = chaves_lancadas_painel.union(chaves_lancadas_titulos)

        # 5. TRATAR PEDIDOS E MAPEAMENTO LOCAL DA OBRA
        df_relacao = carregar_df(file_relacao)
        col_obra_rel = 'Cód. obra' if 'Cód. obra' in df_relacao.columns else df_relacao.columns[0]
        df_relacao['Obra_Clean'] = df_relacao[col_obra_rel].apply(limpar_cod)
        df_relacao_obra = df_relacao[df_relacao['Obra_Clean'] == cod_obra_alvo].copy()

        df_relacao_obra['Cod_Forn'] = df_relacao_obra['Cód. fornecedor'].apply(limpar_cod)
        df_relacao_obra['CNPJ_Forn'] = df_relacao_obra['Cod_Forn'].map(mapa_cod_to_cnpj).fillna("")

        # CNPJs com histórico de Pedidos na Obra
        cnpjs_obra_pedidos = set(df_relacao_obra['CNPJ_Forn'].dropna().unique())
        cnpjs_obra_painel = set(df_painel[df_painel['Cod_Obra_Clean'] == cod_obra_alvo]['CNPJ_Painel'].dropna().unique())
        cnpjs_com_historico_na_obra = cnpjs_obra_pedidos.union(cnpjs_obra_painel)

        # --- ABA 1: PAINEL ---
        resumo_painel = pd.merge(
            df_nf, 
            df_painel[['chave_p', 'N° da Nota fiscal']].drop_duplicates('chave_p'), 
            left_on='chave_unica', 
            right_on='chave_p', 
            how='left'
        )

        def definir_status_painel(r):
            # 1. Se Título/Nota fiscal avulsa estiver preenchido
            col_avulsa = 'Título/Nota fiscal avulsa' if 'Título/Nota fiscal avulsa' in r else None
            if col_avulsa and pd.notna(r[col_avulsa]) and str(r[col_avulsa]).strip() != "":
                return "✅ NF Lançada"
            # 2. Se achou no Global
            if r['chave_unica'] in chaves_lancadas_global:
                return "✅ NF Lançada"
            # 3. Regra restrita à OBRA selecionada
            if r['CNPJ emitente'] in cnpjs_com_historico_na_obra:
                return "⚠️ Para Verificação"
            return "❌ Sem Histórico"

        resumo_painel['Status'] = resumo_painel.apply(definir_status_painel, axis=1)

        # --- ABA 2: PEDIDOS ---
        peds_painel = df_painel[df_painel['Cod_Obra_Clean'] == cod_obra_alvo].groupby('CNPJ_Painel')['N° do Pedido'].apply(lambda x: list(x.dropna().astype(str))).to_dict()
        peds_rel = df_relacao_obra.groupby('CNPJ_Forn')['Nº do pedido'].apply(lambda x: list(x.dropna().astype(str))).to_dict()

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

        # Baixa adicional com a planilha TITULO na Aba Contrato
        titulos_map = df_titulos[['chave_t', 'Documento']].drop_duplicates('chave_t').set_index('chave_t')['Documento'].to_dict() if not df_titulos.empty else {}

        def definir_status_contrato(r):
            if r['chave_unica'] in chaves_lancadas_titulos or r['Status_Ped'] == "✅ Resolvido Painel":
                return "✅ NF Lançada"
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

        # --- ABA 4: EM ABERTO (MANTIDA) ---
        peds_aberto = []
        for _, row in df_painel[df_painel['Cod_Obra_Clean'] == cod_obra_alvo].iterrows():
            st_p = str(row['Situação do pedido']).strip() if pd.notna(row['Situação do pedido']) else ""
            if "cancelado" in st_p.lower(): continue
            ped_n = limpar_cod(row['N° do Pedido'])
            dt_nf = row.get('Data Emissão NF', None)
            num_nf = row.get('N° NF', None)
            if ped_n and pd.isna(dt_nf) and pd.isna(num_nf):
                dt_c = pd.to_datetime(row['Data do pedido'], dayfirst=True, errors='coerce')
                peds_aberto.append({
                    'Pedido': ped_n,
                    'Data Confecção': dt_c,
                    'Fornecedor': row['Fornecedor'],
                    'Status Pedido': st_p
                })

        for _, row in df_relacao_obra.iterrows():
            st_p = str(row['Status entrega']).strip() if pd.notna(row['Status entrega']) else ""
            if "cancelado" in st_p.lower(): continue
            ped_n = limpar_cod(row['Nº do pedido'])
            dt_ent = row.get('Data Entrada NF', None)
            if ped_n and pd.isna(dt_ent):
                dt_c = pd.to_datetime(row['Data pedido'], dayfirst=True, errors='coerce')
                peds_aberto.append({
                    'Pedido': ped_n,
                    'Data Confecção': dt_c,
                    'Fornecedor': row['Fornecedor'],
                    'Status Pedido': st_p
                })

        if peds_aberto:
            df_aberto = pd.DataFrame(peds_aberto).dropna(subset=['Pedido']).groupby('Pedido').agg({
                'Data Confecção': 'min',
                'Fornecedor': 'first',
                'Status Pedido': 'first'
            }).reset_index()
            df_aberto['Dias em aberto'] = (hoje - df_aberto['Data Confecção']).dt.days.fillna(0).astype(int)
            df_aberto['Data Confecção'] = df_aberto['Data Confecção'].dt.strftime('%d/%m/%Y')
            aba4_final = df_aberto[['Pedido', 'Data Confecção', 'Fornecedor', 'Dias em aberto', 'Status Pedido']].sort_values(by='Dias em aberto', ascending=False)
        else:
            aba4_final = pd.DataFrame(columns=['Pedido', 'Data Confecção', 'Fornecedor', 'Dias em aberto', 'Status Pedido'])

        # --- EXPORTAÇÃO FINAL ---
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
