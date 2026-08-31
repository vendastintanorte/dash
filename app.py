import streamlit as st
import pandas as pd
from datetime import datetime
import os

# --- CONFIGURAÇÃO DE PÁGINA ---
st.set_page_config(page_title="Dashboard Vendas", page_icon="⚡", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #f7f9fc; font-family: 'Segoe UI', sans-serif; }
    .main-title { color: #e51e25; font-weight: 900; font-size: 2.8rem; margin-bottom: 0px; text-transform: uppercase; }
    .sub-title { color: #f4ab13; font-size: 1.4rem; font-weight: 700; margin-top: -10px; margin-bottom: 30px; text-transform: uppercase; }
    [data-testid="stMetricValue"] { font-size: 2.2rem !important; font-weight: 900 !important; color: #2b2b2b !important; }
    [data-testid="stMetricLabel"] { font-size: 1rem !important; font-weight: 700 !important; color: #707070 !important; text-transform: uppercase; }
    .section-header { background: linear-gradient(90deg, #f4ab13 0%, #ffc547 100%); padding: 10px 20px; border-radius: 8px; color: #2b2b2b; font-weight: 900; font-size: 1.2rem; margin-top: 30px; margin-bottom: 20px; text-transform: uppercase; }
    .section-header-red { background: linear-gradient(90deg, #e51e25 0%, #ff4b4b 100%); padding: 10px 20px; border-radius: 8px; color: #ffffff; font-weight: 900; font-size: 1.2rem; margin-top: 30px; margin-bottom: 20px; text-transform: uppercase; }
    .alert-box { background-color: #fff3cd; border-left: 5px solid #ffc107; padding: 15px; border-radius: 5px; font-weight: 600; color: #856404; }
    .cat-destaque { background-color: #2b2b2b; color: #f4ab13; padding: 5px 15px; border-radius: 20px; font-size: 1.2rem; font-weight: 800; display: inline-block; margin-bottom: 15px;}
    </style>
""", unsafe_allow_html=True)

PASTA_DADOS = "dados_atuais"
if not os.path.exists(PASTA_DADOS): os.makedirs(PASTA_DADOS)
ARQ_VENDAS = os.path.join(PASTA_DADOS, "vendas.xlsx")
ARQ_RECEBER = os.path.join(PASTA_DADOS, "receber.csv")

# MATRIZ DE CAMPANHAS OFICIAL
CAMPANHAS_MAP = {
    "1. INFINITO / TNT EXCLUSIVE": {"Rebates": "TNT Exclusive", "Camp1": "Conexão Suvinil", "Camp2": "", "Camp3": "", "Camp4": ""},
    "1.1 INFINITO": {"Rebates": "", "Camp1": "Conexão Suvinil", "Camp2": "Vamos Juntos - 1 vaga", "Camp3": "", "Camp4": "Big Fish"},
    "2. DIAMANTE": {"Rebates": "", "Camp1": "Conexão Suvinil", "Camp2": "Vamos juntos - 2 vaga", "Camp3": "Stock Car - 1 vaga", "Camp4": "Big Fish"},
    "3. PLATINUM": {"Rebates": "", "Camp1": "", "Camp2": "Vamos Juntos - 1 vaga", "Camp3": "Stock Car - 2 vaga", "Camp4": "Big Fish"},
    "4. SAFIRA": {"Rebates": "", "Camp1": "", "Camp2": "Vamos Juntos - 1 vaga", "Camp3": "Stock Car - 1 vaga", "Camp4": "Compre e Ganhe"},
    "5. ESMERALDA": {"Rebates": "", "Camp1": "", "Camp2": "Vamos Juntos - 1 vaga", "Camp3": "Stock Car - 1 vaga", "Camp4": "Compre e Ganhe"},
    "6. QUARTZO": {"Rebates": "", "Camp1": "", "Camp2": "Vamos Juntos - 1 vaga", "Camp3": "", "Camp4": "Compre e Ganhe"},
}

st.sidebar.title("Navegação")
aba_selecionada = st.sidebar.radio("Ir para:", ["🔍 Consulta de Clientes", "⚙️ Área do Administrador"])

# ==========================================
# ÁREA DO ADMINISTRADOR
# ==========================================
if aba_selecionada == "⚙️ Área do Administrador":
    st.markdown('<div class="main-title">ÁREA ADMINISTRATIVA</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Atualização Diária de Bases</div>', unsafe_allow_html=True)
    
    senha = st.text_input("Senha de administrador:", type="password")
    if senha == "admin123":
        up_vendas = st.file_uploader("1. Substituir Base de Vendas (Excel)", type=["xlsx"])
        up_receber = st.file_uploader("2. Substituir Contas a Receber (CSV)", type=["csv"])
        if st.button("💾 Salvar Novas Bases"):
            if up_vendas and up_receber:
                try:
                    with open(ARQ_VENDAS, "wb") as f: f.write(up_vendas.getbuffer())
                    with open(ARQ_RECEBER, "wb") as f: f.write(up_receber.getbuffer())
                    st.success("✅ Bases salvas com sucesso! O sistema recarregará os dados instantaneamente.")
                    st.cache_data.clear() # Limpa o cache antigo para forçar leitura nova
                except Exception as e:
                    st.error(f"Erro ao salvar arquivos: {e}")
            else:
                st.error("Por favor, faça o upload de ambos os arquivos.")
    elif senha != "":
        st.error("Senha incorreta.")

# ==========================================
# VISÃO DO VENDEDOR (CONSULTA OTIMIZADA)
# ==========================================
elif aba_selecionada == "🔍 Consulta de Clientes":
    st.markdown('<div class="main-title">DASHBOARD VENDAS</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">INTELIGÊNCIA COMERCIAL EM CAMPO</div>', unsafe_allow_html=True)

    if not os.path.exists(ARQ_VENDAS) or not os.path.exists(ARQ_RECEBER):
        st.warning("⏳ **Atenção:** As bases de dados ainda não foram carregadas. Vá na **Área do Administrador** (menu ao lado), insira a senha e faça a carga inicial.")
        st.stop()

    # CACHE INTELIGENTE: Carrega o arquivo pesado uma única vez na RAM
    @st.cache_data(show_spinner="Carregando e otimizando base de dados volumosa...")
    def carregar_dados_otimizado():
        df_v = pd.read_excel(ARQ_VENDAS, sheet_name=0)
        df_r = pd.read_csv(ARQ_RECEBER, encoding="latin1", sep=None, engine="python")
        
        df_v["DATA EMISSÃO"] = pd.to_datetime(df_v["DATA EMISSÃO"], errors="coerce")
        df_v["ANO"] = df_v["DATA EMISSÃO"].dt.year
        df_v["MES"] = df_v["DATA EMISSÃO"].dt.month
        df_v["Grupo de Cliente"] = df_v["Grupo de Cliente"].fillna(df_v["CLIENTE"]).astype(str).str.strip()
        df_v["FABRICANTE_LAVADO"] = df_v["FABRICANTE"].astype(str).str.strip().str.upper()
        
        df_v["VENDALITROS"] = pd.to_numeric(df_v["VENDALITROS"], errors="coerce").fillna(0)
        df_v["VALORTOTAL"] = pd.to_numeric(df_v["VALORTOTAL"], errors="coerce").fillna(0)
        
        if "VALOR EMABERTO" in df_r.columns:
            df_r["VALOR_NUM"] = df_r["VALOR EMABERTO"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
            df_r["VALOR_NUM"] = pd.to_numeric(df_r["VALOR_NUM"], errors="coerce").fillna(0)
            
        return df_v, df_r

    df_vendas, df_receber = carregar_dados_otimizado()

    HOJE = datetime.now()
    ANO_ATUAL = HOJE.year
    ANO_ANTERIOR = ANO_ATUAL - 1
    MES_ATUAL = HOJE.month

    # --- FILTROS ---
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Buscar Cliente")
    
    cidades = sorted(df_vendas["CIDADE"].dropna().astype(str).str.strip().unique())
    cidade_sel = st.sidebar.selectbox("1. Filtrar por Cidade", ["Todas"] + cidades)
    
    df_filt = df_vendas[df_vendas["CIDADE"].astype(str).str.strip() == cidade_sel] if cidade_sel != "Todas" else df_vendas
    
    grupos_disponiveis = sorted(df_filt["Grupo de Cliente"].dropna().unique())
    if not grupos_disponiveis:
        st.error("Nenhum cliente encontrado nesta cidade.")
        st.stop()
        
    grupo_escolhido = st.sidebar.selectbox("2. Selecione a Rede/Cliente:", grupos_disponiveis)

    # --- PROCESSAMENTO DO GRUPO ---
    df_grupo = df_vendas[df_vendas["Grupo de Cliente"] == grupo_escolhido]
    
    df_atual = df_grupo[(df_grupo["ANO"] == ANO_ATUAL) & (df_grupo["MES"] <= MES_ATUAL)]
    df_anterior = df_grupo[(df_grupo["ANO"] == ANO_ANTERIOR) & (df_grupo["MES"] <= MES_ATUAL)]

    info_grupo = df_grupo.iloc[0]
    categoria_grupo = str(info_grupo.get("CATEGORIA", "Sem Categoria")).strip()

    st.markdown(f'<div class="cat-destaque">🏆 Categoria: {categoria_grupo}</div>', unsafe_allow_html=True)
    st.markdown(f"<h2>{grupo_escolhido}</h2>", unsafe_allow_html=True)

    # ==========================================
    # 1. SUVINIL + SHERWIN (Coluna J: VENDALITROS)
    # ==========================================
    st.markdown(f'<div class="section-header">PERFORMANCE PRINCIPAL (SUVINIL + SHERWIN) - Acumulado Jan a {MES_ATUAL:02d}/{ANO_ATUAL}</div>', unsafe_allow_html=True)
    
    fab_principais = df_grupo[df_grupo["FABRICANTE_LAVADO"].str.contains("SUVINIL|SHERWIN", na=False)]
    L_atual = df_atual[df_atual["FABRICANTE_LAVADO"].str.contains("SUVINIL|SHERWIN", na=False)]["VENDALITROS"].sum()
    L_ant = df_anterior[df_anterior["FABRICANTE_LAVADO"].str.contains("SUVINIL|SHERWIN", na=False)]["VENDALITROS"].sum()

    ultima_compra = fab_principais["DATA EMISSÃO"].max()
    dias_inativo = (HOJE - ultima_compra).days if pd.notna(ultima_compra) else 999
    meses_inativo = dias_inativo // 30
    status_inat = "N/A" if pd.isna(ultima_compra) else (f"⚠️ INATIVO ({meses_inativo} meses)" if meses_inativo >= 3 else f"✅ Ativo")

    if L_ant > 0:
        cresc = ((L_atual - L_ant) / L_ant) * 100
        txt_cresc = f"▲ +{cresc:.1f}% (Ano Ant: {L_ant:,.0f} L)".replace(',', '.') if cresc >= 0 else f"▼ {cresc:.1f}% (Ano Ant: {L_ant:,.0f} L)".replace(',', '.')
    else:
        txt_cresc = "Sem base no ano anterior"

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Positivação", "🟢 SIM" if L_atual > 0 else "🔴 NÃO")
    c2.metric(f"Venda Litros ({ANO_ATUAL})", f"{L_atual:,.0f} L".replace(',', '.'), txt_cresc)
    c3.metric("Pulverização", "META OK" if L_atual >= 50 else f"Faltam {50 - L_atual:,.0f}L".replace(',', '.'), "Alvo: 50 Litros")
    c4.metric("Status Ciclo", status_inat, "Tempo de recompra")

    # ==========================================
    # 2. MIX BÁSICO & HIERARQUIA DE PRODUTOS
    # ==========================================
    st.markdown('<div class="section-header">MIX BÁSICO & HIERARQUIA DE PRODUTOS</div>', unsafe_allow_html=True)
    
    cli_suv_sher = df_atual[df_atual["FABRICANTE_LAVADO"].str.contains("SUVINIL|SHERWIN", na=False)]
    mix_comprado = cli_suv_sher["MIX\nBASICO"].dropna().astype(str).str.upper().unique()
    faltam = [f for f, tem in {"ALVENARIA": any("ALVENARIA" in m for m in mix_comprado), "COMPLEMENTOS": any("COMPLEMENTO" in m for m in mix_comprado), "ESMALTES E VERNIZES": any("ESM" in m for m in mix_comprado)}.items() if not tem]
    
    if not faltam: st.success("🏆 **Mix Básico Completo!** O cliente comprou Alvenaria, Complementos e Esmaltes este ano.")
    else: st.markdown(f'<div class="alert-box">⚠️ FOCO DE VENDA: Falta positivar no Mix Básico este ano: <b>{", ".join(faltam)}</b>.</div>', unsafe_allow_html=True)

    todas_hierarquias = df_vendas[df_vendas["FABRICANTE_LAVADO"].str.contains("SUVINIL|SHERWIN", na=False)]["hierarquia Agrupada"].dropna().unique()
    hierarquias_compradas = cli_suv_sher["hierarquia Agrupada"].dropna().unique()
    hierarquias_faltantes = [h for h in todas_hierarquias if h not in hierarquias_compradas]
    
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        st.markdown(f"**✅ Linhas Já Compradas ({len(hierarquias_compradas)}):**")
        st.dataframe(pd.DataFrame(hierarquias_compradas, columns=["PRODUTOS COMPRADOS"]), use_container_width=True, height=200)
            
    with col_h2:
        st.markdown(f"**❌ Oportunidades - Não Compradas ({len(hierarquias_faltantes)}):**")
        st.dataframe(pd.DataFrame(hierarquias_faltantes, columns=["AÇÕES DE VENDA (FALTANTES)"]), use_container_width=True, height=200)

    # ==========================================
    # 3. OUTROS FORNECEDORES (Amais/Farben = Col J | Adere/Condor = Col L)
    # ==========================================
    st.markdown('<div class="section-header">PERFORMANCE DE MARCAS COMPLEMENTARES</div>', unsafe_allow_html=True)
    
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    
    v_amais_at = df_atual[df_atual["FABRICANTE_LAVADO"].str.contains("AMAIS", na=False)]["VENDALITROS"].sum()
    v_amais_ant = df_anterior[df_anterior["FABRICANTE_LAVADO"].str.contains("AMAIS", na=False)]["VENDALITROS"].sum()
    dif_amais = f"{(((v_amais_at - v_amais_ant) / v_amais_ant) * 100):+.1f}%" if v_amais_ant > 0 else "Sem base"
    with col_f1:
        st.metric("Amais (Litros)", f"{v_amais_at:,.0f} L".replace(',', '.') if v_amais_at > 0 else "-", dif_amais if v_amais_at > 0 else None)

    v_farb_at = df_atual[df_atual["FABRICANTE_LAVADO"].str.contains("FARBEN", na=False)]["VENDALITROS"].sum()
    v_farb_ant = df_anterior[df_anterior["FABRICANTE_LAVADO"].str.contains("FARBEN", na=False)]["VENDALITROS"].sum()
    dif_farb = f"{(((v_farb_at - v_farb_ant) / v_farb_ant) * 100):+.1f}%" if v_farb_ant > 0 else "Sem base"
    with col_f2:
        st.metric("Farben (Litros)", f"{v_farb_at:,.0f} L".replace(',', '.') if v_farb_at > 0 else "-", dif_farb if v_farb_at > 0 else None)

    v_ad_at = df_atual[df_atual["FABRICANTE_LAVADO"].str.contains("ADERE", na=False)]["VALORTOTAL"].sum()
    v_ad_ant = df_anterior[df_anterior["FABRICANTE_LAVADO"].str.contains("ADERE", na=False)]["VALORTOTAL"].sum()
    dif_ad = f"{(((v_ad_at - v_ad_ant) / v_ad_ant) * 100):+.1f}%" if v_ad_ant > 0 else "Sem base"
    with col_f3:
        st.metric("Adere (Faturamento)", f"R$ {v_ad_at:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if v_ad_at > 0 else "-", dif_ad if v_ad_at > 0 else None)

    v_con_at = df_atual[df_atual["FABRICANTE_LAVADO"].str.contains("CONDOR", na=False)]["VALORTOTAL"].sum()
    v_con_ant = df_anterior[df_anterior["FABRICANTE_LAVADO"].str.contains("CONDOR", na=False)]["VALORTOTAL"].sum()
    dif_con = f"{(((v_con_at - v_con_ant) / v_con_ant) * 100):+.1f}%" if v_con_ant > 0 else "Sem base"
    with col_f4:
        st.metric("Condor (Faturamento)", f"R$ {v_con_at:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if v_con_at > 0 else "-", dif_con if v_con_at > 0 else None)

    # ==========================================
    # 4. CAMPANHAS A OFERTAR
    # ==========================================
    st.markdown('<div class="section-header">BENEFÍCIOS E CAMPANHAS (OFERTE NO BALCÃO)</div>', unsafe_allow_html=True)
    camp = CAMPANHAS_MAP.get(categoria_grupo, {})
    
    cc1, cc2, cc3, cc4, cc5 = st.columns(5)
    def box_campanha(titulo, valor):
        return f"""<div style="background-color: white; padding: 15px; border-radius: 8px; border-top: 4px solid #e51e25; box-shadow: 0 2px 4px rgba(0,0,0,0.05); min-height: 110px;">
        <p style="color: #707070; font-size: 0.8rem; font-weight: 700; margin-bottom: 5px; text-transform: uppercase;">{titulo}</p>
        <p style="color: #2b2b2b; font-size: 1rem; font-weight: 800; line-height: 1.2;">{valor}</p></div>"""
    
    with cc1: st.markdown(box_campanha("Rebates", camp.get('Rebates', '-') or '-'), unsafe_allow_html=True)
    with cc2: st.markdown(box_campanha("Ação 1", camp.get('Camp1', '-') or '-'), unsafe_allow_html=True)
    with cc3: st.markdown(box_campanha("Ação 2", camp.get('Camp2', '-') or '-'), unsafe_allow_html=True)
    with cc4: st.markdown(box_campanha("Ação 3", camp.get('Camp3', '-') or '-'), unsafe_allow_html=True)
    with cc5: st.markdown(box_campanha("Ação 4", camp.get('Camp4', '-') or '-'), unsafe_allow_html=True)

    # ==========================================
    # 5. FINANCEIRO
    # ==========================================
    st.markdown('<div class="section-header-red">SITUAÇÃO FINANCEIRA (BOLETOS EM ABERTO)</div>', unsafe_allow_html=True)
    codigos_grupo = [str(cod).zfill(7) for cod in df_grupo["CÓDIGO CLIENTE"].dropna().astype(int).astype(str).unique()]
    boletos_grupo = df_receber[df_receber["CLIENTE"].astype(str).apply(lambda x: any(cod in x for cod in codigos_grupo)) & (df_receber["VALOR_NUM"] > 0)]

    if boletos_grupo.empty: st.success("✅ Tudo limpo! O grupo não possui boletos vencidos ou pendentes.")
    else:
        st.error(f"🛑 ATENÇÃO: O grupo possui {len(boletos_grupo)} boleto(s) em aberto.")
        st.dataframe(boletos_grupo[["CLIENTE", "DOCUMENTO", "EMISSÃO", "VENCIMENTO", "VALOR EMABERTO", "ATRASO"]].sort_values(by="VENCIMENTO"), use_container_width=True)