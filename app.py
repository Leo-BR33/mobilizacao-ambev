import streamlit as st
import pandas as pd
import re
import urllib.parse

# 1. Configuração da Página para Mobile e Web
st.set_page_config(
    page_title="Mobilização Ambev - Prospects",
    page_icon="🎯",
    layout="centered"
)

# 2. Estilo Visual Ambev
st.markdown("""
    <style>
    .header-box {
        background: linear-gradient(135deg, #001489 0%, #0022cc 100%);
        padding: 20px;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.15);
    }
    .header-box h1 {
        color: #FFFFFF !important;
        font-weight: 800;
        margin: 0;
        font-size: 26px;
    }
    .header-box p {
        color: #FBBF24; /* Amarelo Ambev */
        margin-top: 5px;
        font-size: 14px;
        font-weight: 600;
    }
    .prospect-card {
        background-color: #FFFFFF;
        border-left: 6px solid #F59E0B;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 15px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    .badge-bairro { background-color: #FEF3C7; color: #92400E; padding: 4px 10px; border-radius: 5px; font-weight: bold; font-size: 13px; display: inline-block; }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="header-box">
        <h1>🎯 MOBILIZAÇÃO DE VENDAS</h1>
        <p>Prospecção de Novos Pontos de Venda</p>
    </div>
""", unsafe_allow_html=True)

# --- FUNÇÕES DE APOIO ---
def limpar_apenas_numeros(texto):
    if pd.isna(texto): return ""
    return re.sub(r'[\.\-\/\\ ]', '', str(texto)).strip()

def formatar_latitude(valor):
    if pd.isna(valor) or str(valor).strip() in ['', 'nan', 'None']: return None
    val_str = str(valor).strip().replace(',', '.')
    try:
        val_float = float(val_str)
        if -90 <= val_float <= 90 and '.' in val_str: return val_float
    except ValueError: pass

    s = val_str.replace('.', '').replace(',', '').strip()
    sinal = "-" if s.startswith("-") else ""
    s = s.lstrip("-")
    if len(s) >= 2:
        try: return float(f"{sinal}{s[:1]}.{s[1:]}")
        except ValueError: return None
    return None

def formatar_longitude(valor):
    if pd.isna(valor) or str(valor).strip() in ['', 'nan', 'None']: return None
    val_str = str(valor).strip().replace(',', '.')
    try:
        val_float = float(val_str)
        if -180 <= val_float <= 180 and '.' in val_str: return val_float
    except ValueError: pass

    s = val_str.replace('.', '').replace(',', '').strip()
    sinal = "-" if s.startswith("-") else ""
    s = s.lstrip("-")
    if len(s) >= 3:
        try: return float(f"{sinal}{s[:2]}.{s[2:]}")
        except ValueError: return None
    return None

# --- LEITURA DA PLANILHA ---
@st.cache_data(ttl=300)
def carregar_prospects():
    try:
        nome_arquivo = 'base_prospects.xlsx'
        try:
            excel_file = pd.ExcelFile(nome_arquivo, engine='openpyxl')
        except FileNotFoundError:
            nome_arquivo = 'base_clientes.xlsx'
            excel_file = pd.ExcelFile(nome_arquivo, engine='openpyxl')

        # Lê aba 'Prospects' ou a primeira disponível
        aba = 'Prospects' if 'Prospects' in excel_file.sheet_names else excel_file.sheet_names[0]
        data = pd.read_excel(excel_file, sheet_name=aba)

        data.columns = [str(c).strip() for c in data.columns]

        # Mapeia colunas de nomes alternativos
        for col in data.columns:
            c_clean = col.lower().replace('_', '').replace(' ', '')
            if c_clean in ['lat', 'latitude']: data.rename(columns={col: 'Latitude'}, inplace=True)
            elif c_clean in ['long', 'lng', 'longitude']: data.rename(columns={col: 'Longitude'}, inplace=True)
            elif c_clean in ['bairro', 'bairros']: data.rename(columns={col: 'Bairro'}, inplace=True)
            elif c_clean in ['nomefantasia', 'fantasia', 'nome', 'pdv', 'nomedopdv']: data.rename(columns={col: 'Nome_Fantasia'}, inplace=True)

        for col_esperada in ['Nome_Fantasia', 'Bairro', 'Cidade', 'Endereco', 'Telefone']:
            if col_esperada in data.columns: data[col_esperada] = data[col_esperada].astype(str)
            else: data[col_esperada] = ""

        # Coordenadas numéricas limpas
        if 'Latitude' in data.columns: data['lat'] = data['Latitude'].apply(formatar_latitude)
        else: data['lat'] = None

        if 'Longitude' in data.columns: data['lon'] = data['Longitude'].apply(formatar_longitude)
        else: data['lon'] = None

        return data
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {e}")
        return pd.DataFrame()

df = carregar_prospects()

if df.empty:
    st.info("👋 Suba a planilha 'base_prospects.xlsx' no seu repositório do GitHub para começar a mobilização.")
    st.stop()

# --- ABAS DA MOBILIZAÇÃO ---
tab_lista, tab_mapa = st.tabs(["📍 Buscar por Bairro", "🗺️ Mapa Geral de Todos os PDVs"])

# ==============================================================================
# ABA 1: CONSULTA EXCLUSIVA POR BAIRRO
# ==============================================================================
with tab_lista:
    st.markdown("### 🔍 Filtrar Alvos por Bairro")

    bairros_validos = sorted([
        b for b in df['Bairro'].unique() 
        if str(b).strip() and str(b).lower() not in ['nan', 'none', '']
    ])
    opcoes_bairros = ["Todos os Bairros"] + bairros_validos

    col1, col2 = st.columns([2, 2])
    with col1:
        bairro_escolhido = st.selectbox("Selecione o Bairro:", opcoes_bairros)
    with col2:
        busca_nome = st.text_input("Buscar por Nome:", placeholder="Ex: Galeteria, Bar do Zé...")

    # Filtragem
    df_filtrado = df.copy()
    if bairro_escolhido != "Todos os Bairros":
        df_filtrado = df_filtrado[df_filtrado['Bairro'] == bairro_escolhido]

    if busca_nome:
        df_filtrado = df_filtrado[df_filtrado['Nome_Fantasia'].str.contains(busca_nome, case=False, na=False)]

    st.caption(f"🎯 PDVs encontrados: **{len(df_filtrado)}**")

    if not df_filtrado.empty:
        for idx, row in df_filtrado.iterrows():
            nome_pdv = row['Nome_Fantasia'].strip()
            if not nome_pdv or nome_pdv.lower() in ['nan', 'none']:
                nome_pdv = "PDV Prospect"

            bairro = row['Bairro'].strip() or "Bairro não informado"
            cidade = row['Cidade'].strip()
            endereco = row['Endereco'].strip() or "Endereço não informado"

            # Telefone e WhatsApp
            tel = row['Telefone'].replace('.0', '').strip()
            tel_clean = limpar_apenas_numeros(tel)
            tem_tel = bool(tel_clean and tel.lower() not in ['nan', 'none', 'não informado'])
            tel_link = f"https://wa.me/55{tel_clean}" if tem_tel else "#"

            # Geolocalização
            p_lat = row['lat']
            p_lng = row['lon']
            tem_gps = (p_lat is not None) and (p_lng is not None)

            if tem_gps:
                maps_url = f"https://www.google.com/maps/search/?api=1&query={p_lat},{p_lng}"
                map_embed_url = f"https://maps.google.com/maps?q=loc:{p_lat},{p_lng}&z=17&output=embed"
                pode_ver_mapa = True
            elif endereco and endereco != 'Endereço não informado':
                busca_end = f"{endereco}, {bairro}, {cidade}".strip().strip(',')
                q_enc = urllib.parse.quote(busca_end)
                maps_url = f"https://www.google.com/maps/search/?api=1&query={q_enc}"
                map_embed_url = f"https://maps.google.com/maps?q={q_enc}&z=17&output=embed"
                pode_ver_mapa = True
            else:
                maps_url = "#"
                map_embed_url = None
                pode_ver_mapa = False

            st.markdown(f"""
            <div class="prospect-card">
                <h3 style="margin-top:0; color:#001489;">🏬 {nome_pdv}</h3>
                <div style="margin-top: 8px;">
                    <span class="badge-bairro">📍 {bairro}</span>
                </div>
                <div style="margin-top: 8px; font-size: 14px; color: #374151;">
                    <strong>Endereço:</strong> {endereco} {f'- {cidade}' if cidade else ''}<br>
                    <strong>Telefone:</strong> {tel if tem_tel else 'Não informado'}
                </div>
            </div>
            """, unsafe_allow_html=True)

            c_btn1, c_btn2 = st.columns(2)
            with c_btn1:
                if tem_tel:
                    st.link_button("💬 Chamar no WhatsApp", tel_link, use_container_width=True, key=f"wpp_{idx}")
                else:
                    st.button("💬 Sem Telefone", disabled=True, use_container_width=True, key=f"wpp_d_{idx}")
            with c_btn2:
                if pode_ver_mapa:
                    st.link_button("📍 Abrir no GPS / Maps", maps_url, use_container_width=True, key=f"gps_{idx}")
                else:
                    st.button("📍 Sem Localização", disabled=True, use_container_width=True, key=f"gps_d_{idx}")

            if pode_ver_mapa and map_embed_url:
                st.components.v1.iframe(map_embed_url, height=220)

            st.markdown("---")
    else:
        st.warning("Nenhum PDV encontrado para os filtros selecionados.")

# ==============================================================================
# ABA 2: MAPA GERAL E VISÃO DE CLIENTES PRÓXIMOS
# ==============================================================================
with tab_mapa:
    st.markdown("### 🗺️ Mapa Panorâmico da Mobilização")
    st.write("Veja onde estão concentrados todos os alvos e trace rotas pelo Google Maps.")

    col_m1, col_m2 = st.columns([2, 2])
    with col_m1:
        bairro_mapa = st.selectbox("Filtrar Bairro no Mapa:", opcoes_bairros, key="sb_mapa_geral")
    with col_m2:
        st.caption("💡 *Dica:* Aproxime ou afaste o mapa com os dedos no celular.")

    df_mapa = df.dropna(subset=['lat', 'lon']).copy()
    if bairro_mapa != "Todos os Bairros":
        df_mapa = df_mapa[df_mapa['Bairro'] == bairro_mapa]

    st.info(f"📍 **{len(df_mapa)}** PDVs geolocalizados exibidos no mapa.")

    if not df_mapa.empty:
        st.map(
            df_mapa,
            latitude='lat',
            longitude='lon',
            size=35,
            color="#F59E0B",
            use_container_width=True
        )

        st.markdown("#### 🚗 Traçar Rota Rápida no Google Maps")
        for m_idx, m_row in df_mapa.iterrows():
            m_nome = m_row['Nome_Fantasia'].strip() or "PDV Prospect"
            m_bairro = m_row['Bairro'].strip()
            m_end = m_row['Endereco'].strip()
            rota_url = f"https://www.google.com/maps/dir/?api=1&destination={m_row['lat']},{m_row['lon']}"

            col_pdv, col_rota = st.columns([3, 1])
            with col_pdv:
                st.markdown(f"**{m_nome}** ({m_bairro})<br><small style='color:#64748B;'>{m_end}</small>", unsafe_allow_html=True)
            with col_rota:
                st.link_button("🚗 Ir até o PDV", rota_url, use_container_width=True, key=f"rota_{m_idx}")
            st.markdown("<hr style='margin: 4px 0; border: none; border-top: 1px solid #E5E7EB;'>", unsafe_allow_html=True)
    else:
        st.warning("Nenhum PDV com coordenadas válidas para esse filtro.")
