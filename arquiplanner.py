# ==============================================================================
# ARQUIPLANNER 4.0 - GESTÃO INTEGRADA DE PROJETOS, CRIAÇÕES E AGENDA
# Python + Streamlit | Interface: noite musgo, creme e verde-sálvia + busto em 66.000 pontos
# Arquivos na mesma pasta: arquiplanner.py, integracoes_google.py, hero_nuvem.html
# ==============================================================================
import html
import io
import json
import os
import urllib.parse
import uuid
from datetime import date, datetime, timedelta

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components

try:
    import integracoes_google as ig
    ERRO_IG = ""
except Exception as _e:  # o app continua abrindo (sem Google) e mostra o motivo na barra lateral
    ERRO_IG = f"{type(_e).__name__}: {_e}"

    class _IGOff:
        TZ = "America/Sao_Paulo"

        def segredo(self, *a, **k):
            return ""

        def google_ativo(self):
            return False

        drive_ativo = google_ativo

        def drive_backup_baixar(self):
            return None

        def backup_drive_agendado(self, *a, **k):
            return None

        def chamar_gemini(self, *a, **k):
            return "⚠️ Módulo integracoes_google.py não carregou: " + ERRO_IG

    ig = _IGOff()

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, RGBColor
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

esc = html.escape  # todo texto digitado que vai para HTML passa por aqui

# ------------------------------------------------------------------------------
# 1. PÁGINA E ESTILO (paleta e fontes do hero)
# ------------------------------------------------------------------------------
st.set_page_config(page_title="ArquiPlanner | Arquitetura & Gestão", page_icon="📐",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@font-face{font-family:"Geist";src:url(https://fonts.gstatic.com/s/geist/v5/gyByhwUxId8gMEwcGFU.woff2) format("woff2");font-weight:100 900;font-display:swap}
@font-face{font-family:"Geist Mono";src:url(https://fonts.gstatic.com/s/geistmono/v6/or3nQ6H-1_WfwkMZI_qYFrcdmg.woff2) format("woff2");font-weight:100 900;font-display:swap}
@font-face{font-family:"Instrument Serif";src:url(https://fonts.gstatic.com/s/instrumentserif/v5/jizBRFtNs2ka5fXjeivQ4LroWlx-6zUTjg.woff2) format("woff2");font-weight:400;font-display:swap}
:root{--night:#0f110d;--paper:#f0e5c8;--dots:#c9d4b8;--tech:#7fa8a3;--muted:#7f8a6a;--surface:#151912;--surface2:#1b2016;--line:#2a3024;
--sans:"Geist",ui-sans-serif,system-ui,sans-serif;--mono:"Geist Mono",ui-monospace,SFMono-Regular,Menlo,monospace;--serif:"Instrument Serif",Georgia,serif}
@keyframes fadeInUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
@keyframes pulseSage{0%{box-shadow:0 0 0 0 rgba(127,168,163,.45)}70%{box-shadow:0 0 0 8px rgba(127,168,163,0)}100%{box-shadow:0 0 0 0 rgba(127,168,163,0)}}
.stApp{background:var(--night);color:var(--paper);font-family:var(--sans);-webkit-font-smoothing:antialiased}
.stApp::after{content:"";position:fixed;inset:0;pointer-events:none;z-index:999;filter:invert(1);opacity:.05;
background:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='240' height='240'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='.85' numOctaves='3' stitchTiles='stitch'/%3E%3CfeColorMatrix values='0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1.7 0 0 0 -.6'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)'/%3E%3C/svg%3E")}
header[data-testid="stHeader"]{background:transparent}
h1{font-family:var(--serif)!important;font-weight:400!important;letter-spacing:-.01em;color:var(--paper)}
h2,h3,h4,h5{font-family:var(--sans)!important;font-weight:400!important;letter-spacing:-.01em;color:var(--paper)}
label,p,span,li{color:var(--paper)}
section[data-testid="stSidebar"]{background:#0c0e0a;border-right:1px solid var(--line)}
.sidebar-section-title,.mono-label{font:11px/1 var(--mono);letter-spacing:.12em;text-transform:uppercase;color:var(--muted)!important}
.sidebar-section-title{margin:18px 0 8px;padding-bottom:6px;border-bottom:1px solid var(--line)}
.brand{font:26px/1 var(--serif);color:var(--paper)}
/* botões: creme (igual ao pill do hero); navegação lateral em estilo discreto */
.stButton>button,.stDownloadButton>button,.stFormSubmitButton>button,.stLinkButton>a{background:var(--paper)!important;color:#111!important;border:none!important;border-radius:6px!important;font:500 13px var(--sans)!important;letter-spacing:.01em;padding:.45rem 1rem!important;transition:transform .2s,box-shadow .2s,background .2s!important}
.stButton>button:hover,.stDownloadButton>button:hover,.stFormSubmitButton>button:hover,.stLinkButton>a:hover{background:#fff!important;transform:translateY(-2px);box-shadow:0 10px 24px -12px #000}
.stButton>button p,.stDownloadButton>button p,.stFormSubmitButton>button p,.stLinkButton>a p{color:#111!important}
section[data-testid="stSidebar"] .stButton>button{background:transparent!important;color:var(--paper)!important;border:1px solid var(--line)!important;text-align:left;justify-content:flex-start}
section[data-testid="stSidebar"] .stButton>button p{color:var(--paper)!important}
section[data-testid="stSidebar"] .stButton>button:hover{border-color:var(--dots)!important;background:var(--surface)!important;box-shadow:none}
section[data-testid="stSidebar"] .stButton>button[kind="primary"],section[data-testid="stSidebar"] button[data-testid="stBaseButton-primary"]{background:var(--paper)!important;color:#111!important;border-color:var(--paper)!important}
section[data-testid="stSidebar"] .stButton>button[kind="primary"] p,section[data-testid="stSidebar"] button[data-testid="stBaseButton-primary"] p{color:#111!important}
/* cartões */
.arqui-card{background:var(--surface);border:1px solid var(--line);border-left:4px solid var(--dots);border-radius:12px;padding:16px;margin-bottom:14px;animation:fadeInUp .35s ease-out;transition:transform .2s,border-color .2s,box-shadow .2s}
.arqui-card:hover{transform:translateY(-3px);border-color:var(--dots);box-shadow:0 10px 24px -12px #000}
.metric-card{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:16px 18px;animation:fadeInUp .35s ease-out}
.metric-value{font:400 40px/1.1 var(--serif);color:var(--paper);margin:10px 0 4px}
.metric-detail{font:12px var(--sans);color:var(--muted)!important}
.project-hero{background:linear-gradient(90deg,var(--surface) 0%,var(--surface2) 100%);border:1px solid var(--line);border-left:5px solid var(--dots);border-radius:12px;padding:18px 22px;margin-bottom:20px;animation:fadeInUp .3s ease-out}
.badge-tag{background:rgba(127,168,163,.16);color:var(--dots)!important;border:1px solid var(--tech);padding:2px 9px;border-radius:20px;font:600 11px var(--mono)}
.badge-alta{background:#b5483a;color:#fff!important;padding:2px 8px;border-radius:6px;font:600 11px var(--mono)}
.badge-media{background:#d9b46a;color:#111!important;padding:2px 8px;border-radius:6px;font:600 11px var(--mono)}
.badge-baixa{background:var(--tech);color:#0f110d!important;padding:2px 8px;border-radius:6px;font:600 11px var(--mono)}
.stage-pill{display:inline-block;padding:6px 12px;margin:4px 4px 4px 0;border-radius:8px;font:500 12px var(--sans);background:var(--surface);border:1px solid var(--line);color:var(--muted)!important}
.stage-pill-active{background:var(--paper);border-color:var(--paper);color:#111!important;animation:pulseSage 2s infinite}
.stage-pill-done{background:rgba(127,168,163,.14);border-color:var(--tech);color:var(--tech)!important}
.stTextInput>div>div>input,.stTextArea>div>div>textarea,.stSelectbox>div>div>div,.stNumberInput input,.stDateInput input{background:var(--surface)!important;color:var(--paper)!important;border:1px solid var(--line)!important;border-radius:8px!important}
.stTabs [data-baseweb="tab"]{color:var(--muted)}
.stTabs [aria-selected="true"]{color:var(--paper)}
iframe{border:0;border-radius:12px}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# 2. DADOS: JSON LOCAL (escrita atômica) + BACKUP/RESTAURAÇÃO NO DRIVE
# ------------------------------------------------------------------------------
PASTA_APP = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_DADOS = os.path.join(PASTA_APP, "dados_arquiplanner.json")
ARQUIVO_HERO = os.path.join(PASTA_APP, "hero_nuvem.html")


def parse_data(valor, padrao=None):
    """Converte 'AAAA-MM-DD' em date sem quebrar o app se o dado estiver vazio/corrompido."""
    try:
        return date.fromisoformat(str(valor)[:10])
    except Exception:
        return padrao or date.today()


def dados_iniciais_padrao():
    hoje = date.today().isoformat()
    prazo_1 = (date.today() + timedelta(days=14)).isoformat()
    prazo_2 = (date.today() + timedelta(days=30)).isoformat()
    return {
        "config": {"gemini_api_key": "", "drive_pasta_geral": "https://drive.google.com",
                   "google_calendar_embed": "", "nome_usuario": "Jhonatan Lacerda"},
        "etapas_projeto": ["Briefing & Levantamento", "Estudo Preliminar", "Anteprojeto", "Projeto Executivo",
                           "Compatibilização", "Acompanhamento de Obra", "Concluído"],
        "categorias_projeto": ["Arquitetura Residencial", "Arquitetura Comercial", "Design de Interiores",
                               "Paisagismo", "Projeto Executivo / Terceirizado"],
        "projetos": [{
            "id": "proj_01", "titulo": "Residência Lago Sul - Lote 12", "cliente": "Construtora Horizonte",
            "categoria": "Arquitetura Residencial", "etapa_atual": "Estudo Preliminar", "prioridade": "Alta",
            "data_inicio": hoje, "data_fim": prazo_1, "orcamento": 12500.00, "drive_link": "https://drive.google.com",
            "arquivos_drive": [{"nome": "Levantamento Topográfico.dwg", "url": "https://drive.google.com"},
                               {"nome": "Modelo Base BIM.rvt", "url": "https://drive.google.com"}],
            "anotacoes": "Priorizar ventilação cruzada, integração sala/varanda gourmet e brises na fachada oeste.",
            "checklist": [{"item": "Estudo de insolação e ventos", "concluido": True},
                          {"item": "Planta baixa humanizada", "concluido": False},
                          {"item": "Volumetria 3D inicial", "concluido": False}],
            "briefing": {
                "perfil_cliente": "Casal com 2 filhos, gosta de estética contemporânea, concreto aparente, madeira e iluminação indireta.",
                "terreno_imovel": "Terreno 800m² em declive suave, Brasília - DF.",
                "programa_necessidades": "4 suítes, escritório privativo, sala integrada com cozinha gourmet, piscina aquecida e garagem para 3 carros.",
                "restricoes_normas": "Afastamentos obrigatórios do Código de Edificações do DF; taxa de permeabilidade mínima de 30%.",
                "escopo_reuniao": "Reunião inicial: cliente pediu pé-direito duplo no living, cozinha com ilha central e área de serviço isolada.",
                "briefing_ia": ""},
            "revisoes": [{"id": "rev_01", "codigo": "R00 - Estudo Inicial", "data": hoje, "status": "Em Elaboração",
                          "escopo_basico": "- Aumentar largura da varanda gourmet em 1,20m\n- Reposicionar lavabo próximo ao hall\n- Prever ponto de gás e coifa na ilha",
                          "escopo_aprimorado": ""}],
            "cronograma_etapas": [
                {"etapa": "Briefing & Levantamento", "inicio": hoje, "fim": (date.today() + timedelta(days=3)).isoformat()},
                {"etapa": "Estudo Preliminar", "inicio": (date.today() + timedelta(days=4)).isoformat(), "fim": prazo_1},
                {"etapa": "Anteprojeto", "inicio": (date.today() + timedelta(days=15)).isoformat(), "fim": prazo_2}]
        }],
        "categorias_criacoes": ["Atuais", "Futuras", "Cursos em Andamento"],
        "criacoes": [
            {"id": "cria_01", "titulo": "Plugin Python AutoCAD - Merge de Layers", "categoria": "Atuais", "progresso": 75,
             "link": "https://drive.google.com", "proximo_passo": "Finalizar interface CustomTkinter e testar com blocos dinâmicos.",
             "anotacoes": "Automação via win32com.client para agilizar limpeza de plantas recebidas."},
            {"id": "cria_02", "titulo": "Automação RPA & Excel Avançado com Python", "categoria": "Cursos em Andamento", "progresso": 40,
             "link": "https://www.udemy.com", "proximo_passo": "Concluir módulo de integração com planilhas de quantitativos.",
             "anotacoes": "Aplicar nos orçamentos de obra e cronogramas."}],
        "categorias_agenda": [{"nome": "Trabalho", "cor": "#c9d4b8"}, {"nome": "Entregas", "cor": "#d9b46a"},
                              {"nome": "Esportes", "cor": "#7fa8a3"}, {"nome": "Lazer", "cor": "#f0e5c8"},
                              {"nome": "Reunião de Cliente", "cor": "#b5836a"}],
        "eventos_agenda": [
            {"id": "evt_01", "titulo": "Reunião de Alinhamento - Residência Lago Sul", "categoria": "Reunião de Cliente",
             "data": hoje, "hora": "14:30", "descricao": "Apresentação do estudo preliminar e validação de layout."},
            {"id": "evt_02", "titulo": "Corrida Fórmula 1 / Treino", "categoria": "Esportes",
             "data": (date.today() + timedelta(days=2)).isoformat(), "hora": "10:00", "descricao": "Acompanhar etapa da temporada."}]
    }


def mesclar_padrao(dados, padrao):
    """Garante que chaves novas do código existam em dados antigos."""
    for k, v in padrao.items():
        dados.setdefault(k, v)
    for k, v in padrao["config"].items():
        dados["config"].setdefault(k, v)
    return dados


def carregar_dados():
    padrao = dados_iniciais_padrao()
    if os.path.exists(ARQUIVO_DADOS):
        try:
            with open(ARQUIVO_DADOS, "r", encoding="utf-8") as f:
                return mesclar_padrao(json.load(f), padrao)
        except Exception:
            pass
    do_drive = ig.drive_backup_baixar() if ig.drive_ativo() else None  # disco efêmero: restaura do Drive
    dados = mesclar_padrao(do_drive, padrao) if do_drive else padrao
    salvar_dados(dados)
    return dados


def salvar_dados(dados=None):
    if dados is None:
        dados = st.session_state.db
    tmp = ARQUIVO_DADOS + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    os.replace(tmp, ARQUIVO_DADOS)       # escrita atômica
    ig.backup_drive_agendado(dados)      # backup no Drive (no máx. 1x/min)


if "db" not in st.session_state:
    st.session_state.db = carregar_dados()
db = st.session_state.db

# ------------------------------------------------------------------------------
# 3. FUNÇÕES AUXILIARES: GEMINI, DOCX, LINK DA AGENDA, HERO
# ------------------------------------------------------------------------------
def chamar_gemini_ia(prompt_usuario, contexto_sistema=""):
    return ig.chamar_gemini(prompt_usuario, sistema=contexto_sistema,
                            api_key_fallback=db["config"].get("gemini_api_key", ""))


def _titulo_docx(doc, texto):
    t = doc.add_heading(texto, level=0)
    for run in t.runs:
        run.font.color.rgb = RGBColor(86, 120, 116)
    return t


def gerar_docx_revisao(projeto, revisao):
    doc = Document()
    _titulo_docx(doc, "MEMORIAL DE REVISÃO E ESCOPO TÉCNICO").alignment = WD_ALIGN_PARAGRAPH.LEFT
    sub = doc.add_paragraph(f"Projeto: {projeto['titulo']}  |  Revisão: {revisao['codigo']}")
    sub.runs[0].bold = True
    sub.runs[0].font.size = Pt(12)
    tabela = doc.add_table(rows=4, cols=2)
    tabela.style = "Table Grid"
    infos = [("Cliente / Empresa:", projeto.get("cliente", "-")), ("Categoria / Tipologia:", projeto.get("categoria", "-")),
             ("Etapa Atual:", projeto.get("etapa_atual", "-")),
             ("Data da Revisão / Prazo Final:", f"{revisao.get('data', '-')}  |  Entrega: {projeto.get('data_fim', '-')}")]
    for idx, (chave, valor) in enumerate(infos):
        row = tabela.rows[idx]
        row.cells[0].text = chave
        row.cells[1].text = str(valor)
        row.cells[0].paragraphs[0].runs[0].bold = True
    doc.add_paragraph("")
    doc.add_heading("1. Escopo Técnico e Diretrizes da Revisão", level=1)
    texto_final = revisao.get("escopo_aprimorado", "").strip() or revisao.get("escopo_basico", "").strip()
    for linha in texto_final.split("\n"):
        if linha.strip():
            doc.add_paragraph(linha.strip()).paragraph_format.space_after = Pt(4)
    if projeto.get("checklist"):
        doc.add_heading("2. Checklist de Entregáveis e Pendências", level=1)
        for item in projeto["checklist"]:
            doc.add_paragraph(f"{'[X]' if item['concluido'] else '[  ]'} {item['item']}", style="List Bullet")
    doc.add_paragraph("")
    rodape = doc.add_paragraph(f"Documento gerado via ArquiPlanner | {db['config'].get('nome_usuario', '')}")
    rodape.runs[0].italic = True
    rodape.runs[0].font.size = Pt(9)
    rodape.runs[0].font.color.rgb = RGBColor(120, 120, 120)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def gerar_docx_briefing(projeto):
    doc = Document()
    b = projeto.get("briefing", {})
    _titulo_docx(doc, "BRIEFING ARQUITETÔNICO & PROGRAMA DE NECESSIDADES")
    doc.add_paragraph(f"Projeto: {projeto['titulo']}  |  Cliente: {projeto['cliente']}")
    doc.add_paragraph(f"Categoria: {projeto['categoria']}  |  Data de Início: {projeto['data_inicio']}  |  Prazo: {projeto['data_fim']}")
    if b.get("briefing_ia", "").strip():
        doc.add_heading("Briefing Estruturado (Versão Executiva)", level=1)
        for linha in b["briefing_ia"].split("\n"):
            if linha.strip():
                doc.add_paragraph(linha.strip())
    else:
        for titulo_sec, conteudo in [
            ("1. Perfil do Cliente e Referências Estéticas", b.get("perfil_cliente", "")),
            ("2. Terreno / Imóvel e Localização", b.get("terreno_imovel", "")),
            ("3. Programa de Necessidades", b.get("programa_necessidades", "")),
            ("4. Restrições Técnicas e Normativas", b.get("restricoes_normas", "")),
            ("5. Anotações da Reunião de Escopo", b.get("escopo_reuniao", ""))]:
            doc.add_heading(titulo_sec, level=1)
            doc.add_paragraph(conteudo or "Não preenchido.")
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def gerar_link_google_agenda(titulo, data_str, hora_str="09:00", descricao=""):
    try:
        dt_base = datetime.strptime(f"{data_str} {hora_str}", "%Y-%m-%d %H:%M")
    except ValueError:
        dt_base = datetime.now()
    dt_fim = dt_base + timedelta(hours=1)
    params = urllib.parse.urlencode({
        "action": "TEMPLATE", "text": titulo, "ctz": ig.TZ, "details": descricao,
        "dates": f"{dt_base.strftime('%Y%m%dT%H%M%S')}/{dt_fim.strftime('%Y%m%dT%H%M%S')}"})
    return f"https://calendar.google.com/calendar/render?{params}"


def render_hero(nome, tagline, hint, altura=720):
    """Busto em nuvem de pontos (WebGL/three.js) dentro de um iframe do Streamlit."""
    try:
        with open(ARQUIVO_HERO, "r", encoding="utf-8") as f:
            doc = f.read()
    except FileNotFoundError:
        st.title("ArquiPlanner")
        st.caption("hero_nuvem.html não encontrado na pasta do app.")
        return
    doc = doc.replace("__NOME__", esc(nome)).replace("__TAG__", esc(tagline)).replace("__HINT__", esc(hint))
    components.html(doc, height=altura)


def abrir_projeto(pid):
    """Callback de botão: troca o projeto ativo (e o selectbox) antes do rerun."""
    st.session_state.projeto_selecionado_id = pid
    st.session_state["seletor_global_projeto"] = pid


def card_metrica(rotulo, valor, detalhe=""):
    return (f'<div class="metric-card"><div class="mono-label">{esc(rotulo)}</div>'
            f'<div class="metric-value">{esc(str(valor))}</div><div class="metric-detail">{esc(detalhe)}</div></div>')


# ------------------------------------------------------------------------------
# 4. BARRA LATERAL
# ------------------------------------------------------------------------------
MENU_INICIO = "🏠 Início"
opcoes_projetos = ["📋 Visão Geral & Edição Completa", "📝 Briefing", "🔄 Revisões & Escopo IA (.docx)",
                   "📊 Gráfico de Andamento & Etapas", "🏷️ Categorias & Etapas de Projetos"]
opcoes_criacoes = [f"✨ {c}" for c in db["categorias_criacoes"]] + ["➕ Nova Categoria de Criação"]
opcoes_agenda = ["🗓️ Agenda Google & Categorias"]

if "menu_ativo" not in st.session_state or st.session_state.menu_ativo not in (
        [MENU_INICIO] + opcoes_projetos + opcoes_criacoes + opcoes_agenda):
    st.session_state.menu_ativo = MENU_INICIO


def botao_nav(rotulo):
    ativo = st.session_state.menu_ativo == rotulo
    if st.button(rotulo, key=f"nav_{rotulo}", use_container_width=True, type="primary" if ativo else "secondary"):
        st.session_state.menu_ativo = rotulo
        st.rerun()


with st.sidebar:
    st.markdown('<div class="brand">ArquiPlanner.</div>', unsafe_allow_html=True)
    st.caption("Estúdio de Arquitetura & Gestão")
    botao_nav(MENU_INICIO)
    st.markdown('<div class="sidebar-section-title">Projetos</div>', unsafe_allow_html=True)
    for item in opcoes_projetos:
        botao_nav(item)
    st.markdown('<div class="sidebar-section-title">Criações & Estudos</div>', unsafe_allow_html=True)
    for item in opcoes_criacoes:
        botao_nav(item)
    st.markdown('<div class="sidebar-section-title">Calendário & Rotina</div>', unsafe_allow_html=True)
    for item in opcoes_agenda:
        botao_nav(item)

    st.markdown("---")
    with st.expander("⚙️ Configurações"):
        novo_nome = st.text_input("Nome exibido na tela inicial:", value=db["config"].get("nome_usuario", ""))
        nova_key = st.text_input("Chave API Gemini (se não usar secrets):", value=db["config"].get("gemini_api_key", ""), type="password")
        novo_drive = st.text_input("Link Pasta Raiz Google Drive:", value=db["config"].get("drive_pasta_geral", ""))
        novo_embed = st.text_input("URL Embed Google Agenda:", value=db["config"].get("google_calendar_embed", ""),
                                   help="Cole o link 'Incorporar' das configurações do seu Google Agenda.")
        if st.button("Salvar Configurações", use_container_width=True):
            db["config"].update({"nome_usuario": novo_nome, "gemini_api_key": nova_key,
                                 "drive_pasta_geral": novo_drive, "google_calendar_embed": novo_embed})
            salvar_dados()
            st.success("Configurações salvas!")

    if ERRO_IG:
        st.error(f"integracoes_google.py não carregou. {ERRO_IG}")
    st.caption(f"Agenda: {'🟢 conectada' if ig.google_ativo() else '⚪ não configurada'}  |  "
               f"Drive: {'🟢' if ig.drive_ativo() else '⚪'}  |  "
               f"Gemini: {'🟢' if (ig.segredo('GEMINI_API_KEY') or db['config'].get('gemini_api_key')) else '⚪'}")
    if ig.drive_ativo() and ig.segredo("DRIVE_BACKUP_FILE_ID"):
        if st.button("☁️ Backup agora no Drive", use_container_width=True):
            try:
                ig.drive_backup_enviar(db)
                st.success("Backup enviado!")
            except Exception as e:
                st.error(f"Falha no backup: {e}")
    if st.session_state.get("_erro_backup"):
        st.warning("O último backup automático falhou.")

menu_atual = st.session_state.menu_ativo


# ------------------------------------------------------------------------------
# 5. SELETOR GLOBAL DE PROJETO ATIVO
# ------------------------------------------------------------------------------
def obter_projeto_ativo():
    projetos = db["projetos"]
    if not projetos:
        return None
    nomes = {p["id"]: f"{p['titulo']} — ({p['cliente']})" for p in projetos}
    ids = list(nomes.keys())
    if st.session_state.get("projeto_selecionado_id") not in ids:
        st.session_state.projeto_selecionado_id = ids[0]
    if st.session_state.get("seletor_global_projeto") not in ids:
        st.session_state["seletor_global_projeto"] = st.session_state.projeto_selecionado_id

    col_sel, col_drive, col_fluxo = st.columns([3, 1, 2])
    with col_sel:
        escolhido = st.selectbox("Projeto selecionado:", options=ids, format_func=lambda x: nomes[x], key="seletor_global_projeto")
        st.session_state.projeto_selecionado_id = escolhido
    proj = next(p for p in projetos if p["id"] == escolhido)

    with col_drive:
        st.write("")
        st.write("")
        st.link_button("📂 Abrir Google Drive", proj.get("drive_link") or db["config"].get("drive_pasta_geral", "https://drive.google.com"),
                       use_container_width=True)
    with col_fluxo:
        st.write("")
        st.write("")
        c1, c2, c3 = st.columns(3)
        for col, rotulo, destino, k in [(c1, "📝 Briefing", "📝 Briefing", "jump_briefing"),
                                        (c2, "🔄 Revisões", "🔄 Revisões & Escopo IA (.docx)", "jump_revisoes"),
                                        (c3, "📊 Gráfico", "📊 Gráfico de Andamento & Etapas", "jump_grafico")]:
            if col.button(rotulo, key=k, use_container_width=True):
                st.session_state.menu_ativo = destino
                st.rerun()

    checklist = proj.get("checklist", [])
    feitos = sum(1 for c in checklist if c["concluido"])
    pct = int(feitos / len(checklist) * 100) if checklist else 0
    cls_pri = {"Alta": "alta", "Média": "media"}.get(proj["prioridade"], "baixa")
    st.markdown(f"""
    <div class="project-hero"><div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;">
      <div><span class="badge-tag">{esc(proj['categoria'])}</span>
        <span class="badge-{cls_pri}" style="margin-left:6px;">Prioridade {esc(proj['prioridade'])}</span>
        <h2 style="margin:8px 0 4px 0;">{esc(proj['titulo'])}</h2>
        <p style="color:#7f8a6a;margin:0;font-size:13px;">Cliente: <b style="color:#f0e5c8">{esc(proj['cliente'])}</b> &nbsp;|&nbsp;
        Etapa: <b style="color:#f0e5c8">{esc(proj['etapa_atual'])}</b> &nbsp;|&nbsp; Entrega: <b style="color:#f0e5c8">{esc(proj['data_fim'])}</b>
        &nbsp;|&nbsp; Honorários: <b style="color:#f0e5c8">R$ {float(proj['orcamento']):,.2f}</b></p></div>
      <div style="text-align:right;"><span class="mono-label">Checklist concluído</span>
        <h3 style="margin:6px 0 0 0;color:#c9d4b8;">{pct}% ({feitos}/{len(checklist)})</h3></div>
    </div></div>""", unsafe_allow_html=True)
    return proj


def estilo_plotly(fig, altura):
    fig.update_layout(paper_bgcolor="#0f110d", plot_bgcolor="#151912", font_color="#f0e5c8", height=altura,
                      font_family="Geist, sans-serif")
    return fig


PALETA = ["#c9d4b8", "#7fa8a3", "#f0e5c8", "#d9b46a", "#b5836a"]

# ==============================================================================
# TELA INICIAL: BUSTO EM PONTOS + PAINEL DO DIA
# ==============================================================================
if menu_atual == MENU_INICIO:
    render_hero(db["config"].get("nome_usuario") or "ArquiPlanner",
                "Projetos, criações e agenda num só lugar", "Mova a luz")

    hoje = date.today()
    ativos = [p for p in db["projetos"] if p["etapa_atual"] != "Concluído"]
    prazo7 = sorted([p for p in ativos if hoje <= parse_data(p["data_fim"]) <= hoje + timedelta(days=7)],
                    key=lambda p: p["data_fim"])
    atrasados = [p for p in ativos if parse_data(p["data_fim"]) < hoje]
    eventos_hoje = sorted([e for e in db["eventos_agenda"] if e["data"] == hoje.isoformat()], key=lambda e: e.get("hora", ""))
    criacoes_andamento = [c for c in db["criacoes"] if int(c.get("progresso", 0)) < 100]

    m1, m2, m3, m4 = st.columns(4)
    m1.markdown(card_metrica("Projetos ativos", len(ativos), f"{len(atrasados)} com prazo vencido" if atrasados else "nenhum atrasado"), unsafe_allow_html=True)
    m2.markdown(card_metrica("Entregas em 7 dias", len(prazo7), "próximos prazos"), unsafe_allow_html=True)
    m3.markdown(card_metrica("Compromissos hoje", len(eventos_hoje), hoje.strftime("%d/%m/%Y")), unsafe_allow_html=True)
    m4.markdown(card_metrica("Criações em andamento", len(criacoes_andamento), "estudos, plugins e cursos"), unsafe_allow_html=True)

    st.write("")
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Próximas entregas")
        if prazo7 or atrasados:
            for p in (atrasados + prazo7)[:5]:
                venceu = parse_data(p["data_fim"]) < hoje
                st.markdown(f"""<div class="arqui-card" style="padding:12px;{'border-left-color:#b5483a;' if venceu else ''}">
                    <span class="mono-label">{esc(p['data_fim'])}{' · VENCIDO' if venceu else ''}</span>
                    <h4 style="margin:6px 0 2px 0;">{esc(p['titulo'])}</h4>
                    <p style="font-size:12px;color:#7f8a6a;margin:0;">{esc(p['cliente'])} — {esc(p['etapa_atual'])}</p></div>""", unsafe_allow_html=True)
        else:
            st.info("Nenhuma entrega nos próximos 7 dias.")
    with col_b:
        st.subheader("Hoje na agenda")
        if eventos_hoje:
            for e in eventos_hoje:
                st.markdown(f"""<div class="arqui-card" style="padding:12px;">
                    <span class="mono-label">{esc(e.get('hora', ''))} · {esc(e['categoria'])}</span>
                    <h4 style="margin:6px 0 2px 0;">{esc(e['titulo'])}</h4>
                    <p style="font-size:12px;color:#7f8a6a;margin:0;">{esc(e.get('descricao', ''))}</p></div>""", unsafe_allow_html=True)
        else:
            st.info("Nenhum compromisso para hoje.")

    st.markdown('<div class="mono-label" style="margin:18px 0 8px;">Acesso rápido</div>', unsafe_allow_html=True)
    q1, q2, q3, q4 = st.columns(4)
    for col, rotulo, destino in [(q1, "📋 Projetos", opcoes_projetos[0]), (q2, "📝 Briefing", opcoes_projetos[1]),
                                 (q3, "🔄 Revisões", opcoes_projetos[2]), (q4, "🗓️ Agenda", opcoes_agenda[0])]:
        if col.button(rotulo, key=f"quick_{destino}", use_container_width=True):
            st.session_state.menu_ativo = destino
            st.rerun()

# ==============================================================================
# TELA A: VISÃO GERAL & EDIÇÃO PROFUNDA
# ==============================================================================
elif menu_atual == "📋 Visão Geral & Edição Completa":
    st.title("Projetos")

    with st.expander("➕ Criar Novo Projeto"):
        with st.form("form_criar_projeto", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            p_titulo = c1.text_input("Nome do Projeto *")
            p_cliente = c2.text_input("Cliente / Empresa")
            p_cat = c3.selectbox("Categoria do Projeto", db["categorias_projeto"])
            c4, c5, c6, c7 = st.columns(4)
            p_etapa = c4.selectbox("Etapa Inicial", db["etapas_projeto"])
            p_pri = c5.selectbox("Prioridade", ["Alta", "Média", "Baixa"])
            p_ini = c6.date_input("Data de Início", date.today())
            p_fim = c7.date_input("Data de Entrega", date.today() + timedelta(days=20))
            c8, c9 = st.columns(2)
            p_orc = c8.number_input("Honorários / Valor (R$)", min_value=0.0, step=250.0)
            p_drive = c9.text_input("Link da Pasta no Google Drive", value=db["config"].get("drive_pasta_geral", "https://drive.google.com"))
            p_notas = st.text_area("Anotações Gerais / Diretrizes Iniciais")

            if st.form_submit_button("Salvar Novo Projeto") and p_titulo.strip():
                link_pasta = p_drive
                if ig.drive_ativo() and p_drive.startswith("http") and "drive.google.com/drive" in p_drive:
                    try:  # cria uma subpasta para o projeto dentro da pasta informada
                        link_pasta = ig.drive_criar_pasta(p_titulo.strip(), p_drive)
                    except Exception:
                        link_pasta = p_drive
                novo_id = f"proj_{uuid.uuid4().hex[:6]}"
                db["projetos"].append({
                    "id": novo_id, "titulo": p_titulo.strip(), "cliente": p_cliente.strip() or "Cliente Particular",
                    "categoria": p_cat, "etapa_atual": p_etapa, "prioridade": p_pri,
                    "data_inicio": p_ini.isoformat(), "data_fim": p_fim.isoformat(), "orcamento": p_orc,
                    "drive_link": link_pasta, "arquivos_drive": [], "anotacoes": p_notas, "checklist": [],
                    "briefing": {"perfil_cliente": "", "terreno_imovel": "", "programa_necessidades": "",
                                 "restricoes_normas": "", "escopo_reuniao": p_notas, "briefing_ia": ""},
                    "revisoes": [],
                    "cronograma_etapas": [{"etapa": p_etapa, "inicio": p_ini.isoformat(), "fim": p_fim.isoformat()}]})
                abrir_projeto(novo_id)
                salvar_dados()
                st.success(f"Projeto '{p_titulo}' criado!")
                st.rerun()

    aba_editor, aba_kanban = st.tabs(["✏️ Editor do Projeto Ativo", "🗂️ Quadro Kanban por Etapas"])

    with aba_editor:
        proj = obter_projeto_ativo()
        if proj:
            col_esq, col_dir = st.columns([3, 2])
            with col_esq:
                st.subheader("Dados Principais & Prazos")
                with st.form("form_editar_projeto_completo"):
                    e_titulo = st.text_input("Título do Projeto", value=proj["titulo"])
                    e_cliente = st.text_input("Cliente / Escritório Parceiro", value=proj["cliente"])
                    ec1, ec2, ec3 = st.columns(3)
                    idx_cat = db["categorias_projeto"].index(proj["categoria"]) if proj["categoria"] in db["categorias_projeto"] else 0
                    e_cat = ec1.selectbox("Categoria", db["categorias_projeto"], index=idx_cat)
                    idx_etp = db["etapas_projeto"].index(proj["etapa_atual"]) if proj["etapa_atual"] in db["etapas_projeto"] else 0
                    e_etapa = ec2.selectbox("Etapa Atual", db["etapas_projeto"], index=idx_etp)
                    pris = ["Alta", "Média", "Baixa"]
                    e_pri = ec3.selectbox("Prioridade", pris, index=pris.index(proj["prioridade"]) if proj["prioridade"] in pris else 1)
                    ed1, ed2, ed3 = st.columns(3)
                    e_ini = ed1.date_input("Data de Início", value=parse_data(proj["data_inicio"]))
                    e_fim = ed2.date_input("Prazo Final", value=parse_data(proj["data_fim"]))
                    e_orc = ed3.number_input("Honorários (R$)", value=float(proj["orcamento"]), step=250.0)
                    e_drive = st.text_input("Link Principal da Pasta no Google Drive", value=proj.get("drive_link", ""))
                    e_notas = st.text_area("Anotações Gerais do Projeto", value=proj.get("anotacoes", ""), height=130)
                    if st.form_submit_button("💾 Salvar Alterações do Projeto"):
                        proj.update({"titulo": e_titulo, "cliente": e_cliente, "categoria": e_cat, "etapa_atual": e_etapa,
                                     "prioridade": e_pri, "data_inicio": e_ini.isoformat(), "data_fim": e_fim.isoformat(),
                                     "orcamento": e_orc, "drive_link": e_drive, "anotacoes": e_notas})
                        salvar_dados()
                        st.success("Alterações salvas com sucesso!")
                        st.rerun()

                with st.expander("🗑️ Zona de Perigo (Excluir Projeto)"):
                    confirma = st.checkbox("Tenho certeza que quero excluir este projeto", key=f"conf_del_{proj['id']}")
                    if st.button("Excluir definitivamente", key="del_proj_btn", disabled=not confirma):
                        db["projetos"] = [p for p in db["projetos"] if p["id"] != proj["id"]]
                        salvar_dados()
                        st.warning("Projeto excluído.")
                        st.rerun()

            with col_dir:
                st.subheader("✅ Checklist de Atividades")
                novo_item_chk = st.text_input("Adicionar nova tarefa ao checklist:", key="novo_chk_input")
                if st.button("➕ Adicionar Tarefa", key="btn_add_chk") and novo_item_chk.strip():
                    proj["checklist"].append({"item": novo_item_chk.strip(), "concluido": False})
                    salvar_dados()
                    st.rerun()
                for i, chk in enumerate(proj["checklist"]):
                    c_box, c_txt, c_del = st.columns([0.12, 0.73, 0.15])
                    marcado = c_box.checkbox("Concluído", value=chk["concluido"], key=f"chk_done_{proj['id']}_{i}", label_visibility="collapsed")
                    if marcado != chk["concluido"]:
                        chk["concluido"] = marcado
                        salvar_dados()
                        st.rerun()
                    c_txt.markdown(f"~~{chk['item']}~~" if chk["concluido"] else chk["item"])
                    if c_del.button("✕", key=f"chk_del_{proj['id']}_{i}"):
                        proj["checklist"].pop(i)
                        salvar_dados()
                        st.rerun()

                st.markdown("---")
                st.subheader("📂 Arquivos Vinculados (Google Drive)")
                if ig.drive_ativo() and proj.get("drive_link"):
                    if st.button("🔄 Buscar arquivos da pasta no Drive", key="btn_scan_drive"):
                        try:
                            achados = ig.drive_listar(proj["drive_link"])
                            ja = {a["url"] for a in proj.get("arquivos_drive", [])}
                            novos = [{"nome": f["name"], "url": f["webViewLink"]} for f in achados if f["webViewLink"] not in ja]
                            proj.setdefault("arquivos_drive", []).extend(novos)
                            salvar_dados()
                            st.success(f"{len(novos)} arquivo(s) novo(s) vinculado(s).")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Não consegui ler a pasta (ela foi compartilhada com a conta de serviço?): {e}")
                with st.form("form_add_arquivo_drive", clear_on_submit=True):
                    arq_nome = st.text_input("Nome do Arquivo (ex: Planta Executiva.dwg)")
                    arq_url = st.text_input("Link do Arquivo no Google Drive")
                    if st.form_submit_button("Vincular Arquivo") and arq_nome and arq_url:
                        proj.setdefault("arquivos_drive", []).append({"nome": arq_nome, "url": arq_url})
                        salvar_dados()
                        st.rerun()
                for idx_a, arq in enumerate(proj.get("arquivos_drive", [])):
                    ca1, ca2 = st.columns([0.82, 0.18])
                    ca1.markdown(f"📄 [{arq['nome']}]({arq['url']})")
                    if ca2.button("✕", key=f"del_arq_{proj['id']}_{idx_a}"):
                        proj["arquivos_drive"].pop(idx_a)
                        salvar_dados()
                        st.rerun()
        else:
            st.info("Nenhum projeto cadastrado. Crie o primeiro projeto acima!")

    with aba_kanban:
        etapas = db["etapas_projeto"]
        for ini in range(0, len(etapas), 4):  # no máximo 4 colunas por linha
            cols_k = st.columns(4)
            for j, etapa_nome in enumerate(etapas[ini:ini + 4]):
                with cols_k[j]:
                    st.markdown(f'<div class="mono-label" style="margin:10px 0 8px;">{esc(etapa_nome)}</div>', unsafe_allow_html=True)
                    for p in [p for p in db["projetos"] if p["etapa_atual"] == etapa_nome]:
                        st.markdown(f"""<div class="arqui-card"><span class="badge-tag">{esc(p['categoria'])}</span>
                            <h4 style="margin:8px 0 4px 0;">{esc(p['titulo'])}</h4>
                            <p style="font-size:12px;color:#7f8a6a;margin:0;">{esc(p['cliente'])}</p>
                            <p style="font-size:11px;color:#c9d4b8;margin-top:4px;">Entrega: {esc(p['data_fim'])}</p></div>""", unsafe_allow_html=True)
                        st.button("Abrir Editor", key=f"open_k_{p['id']}", use_container_width=True,
                                  on_click=abrir_projeto, args=(p["id"],))

# ==============================================================================
# TELA B: BRIEFING (GEMINI + .DOCX)
# ==============================================================================
elif menu_atual == "📝 Briefing":
    st.title("Briefing Arquitetônico")
    proj = obter_projeto_ativo()
    if proj:
        b = proj.setdefault("briefing", {"perfil_cliente": "", "terreno_imovel": "", "programa_necessidades": "",
                                         "restricoes_normas": "", "escopo_reuniao": "", "briefing_ia": ""})
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.subheader("1. Levantamento & Diretrizes")
            with st.form("form_briefing_projeto"):
                b_perfil = st.text_area("Perfil do Cliente & Estilo Desejado", value=b.get("perfil_cliente", ""), height=95)
                b_terreno = st.text_area("Terreno / Imóvel, Insolação e Metragem", value=b.get("terreno_imovel", ""), height=95)
                b_prog = st.text_area("Programa de Necessidades (Ambientes e Fluxos)", value=b.get("programa_necessidades", ""), height=110)
                b_rest = st.text_area("Restrições Técnicas, NBRs e Legislação", value=b.get("restricoes_normas", ""), height=95)
                b_escopo = st.text_area("Notas Brutas da Reunião de Briefing", value=b.get("escopo_reuniao", ""), height=110)
                if st.form_submit_button("💾 Salvar Dados do Briefing"):
                    b.update({"perfil_cliente": b_perfil, "terreno_imovel": b_terreno, "programa_necessidades": b_prog,
                              "restricoes_normas": b_rest, "escopo_reuniao": b_escopo})
                    salvar_dados()
                    st.success("Briefing salvo!")
                    st.rerun()
        with col_b2:
            st.subheader("2. Consolidação Executiva com Gemini")
            st.write("A IA organiza suas anotações brutas em um briefing técnico pronto para enviar ao cliente ou anexar no Drive.")
            if st.button("✨ Gerar / Aprimorar Briefing com Gemini", use_container_width=True):
                with st.spinner("O Gemini está estruturando o briefing..."):
                    prompt_b = f"""
                    Atue como um Arquiteto e Urbanista Sênior. Organize os dados abaixo em um Briefing Arquitetônico Executivo claro, técnico e bem estruturado em português (sem usar símbolos excessivos de markdown como ### ou **, use títulos numerados limpos):
                    - Projeto: {proj['titulo']}
                    - Cliente: {proj['cliente']}
                    - Perfil do Cliente: {b.get('perfil_cliente', '')}
                    - Terreno/Imóvel: {b.get('terreno_imovel', '')}
                    - Programa de Necessidades: {b.get('programa_necessidades', '')}
                    - Restrições e Normas: {b.get('restricoes_normas', '')}
                    - Anotações Brutas de Reunião: {b.get('escopo_reuniao', '')}
                    """
                    b["briefing_ia"] = chamar_gemini_ia(prompt_b)
                    salvar_dados()
                    st.rerun()
            briefing_editado = st.text_area("Texto Final do Briefing (Editável):", value=b.get("briefing_ia", ""),
                                            height=380, key=f"txt_briefing_ia_{proj['id']}")
            if st.button("💾 Salvar Texto Final Editado", key="save_briefing_final"):
                b["briefing_ia"] = briefing_editado
                salvar_dados()
                st.success("Texto final atualizado!")
            if HAS_DOCX:
                st.download_button("📥 Baixar Briefing em Word (.docx)", data=gerar_docx_briefing(proj),
                                   file_name=f"Briefing_{proj['titulo'].replace(' ', '_')}.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   use_container_width=True)
            else:
                st.warning("Instale `python-docx` (`pip install python-docx`) para habilitar o download em Word.")

# ==============================================================================
# TELA C: REVISÕES & ESCOPO (GEMINI + .DOCX)
# ==============================================================================
elif menu_atual == "🔄 Revisões & Escopo IA (.docx)":
    st.title("Revisões & Escopo Inteligente")
    proj = obter_projeto_ativo()
    if proj:
        revisoes = proj.setdefault("revisoes", [])
        STATUS = ["Em Elaboração", "Enviado ao Cliente", "Aprovado"]
        with st.expander("➕ Registrar Nova Revisão (ex: R01, R02, Compatibilização)", expanded=(len(revisoes) == 0)):
            with st.form("form_nova_revisao", clear_on_submit=True):
                rc1, rc2, rc3 = st.columns(3)
                r_cod = rc1.text_input("Código / Título da Revisão *", placeholder="Ex: R01 - Ajustes Pós-Reunião")
                r_data = rc2.date_input("Data da Reunião / Revisão", date.today())
                r_status = rc3.selectbox("Status", STATUS)
                r_basico = st.text_area("Escopo Básico / Anotações Rápidas Durante a Reunião:", height=130,
                                        placeholder="Escreva tópicos rápidos anotados na reunião. Depois use a IA para transformar em escopo técnico formal...")
                if st.form_submit_button("Criar Revisão") and r_cod.strip():
                    revisoes.insert(0, {"id": f"rev_{uuid.uuid4().hex[:6]}", "codigo": r_cod.strip(), "data": r_data.isoformat(),
                                        "status": r_status, "escopo_basico": r_basico, "escopo_aprimorado": ""})
                    salvar_dados()
                    st.success("Revisão registrada!")
                    st.rerun()

        if revisoes:
            opcoes_rev = {r["id"]: f"{r['codigo']} ({r['data']}) — [{r['status']}]" for r in revisoes}
            rev_id_sel = st.selectbox("Revisão para editar / processar com Gemini:", options=list(opcoes_rev.keys()),
                                      format_func=lambda x: opcoes_rev[x])
            rev = next(r for r in revisoes if r["id"] == rev_id_sel)
            st.markdown("---")
            col_r1, col_r2 = st.columns(2)
            with col_r1:
                st.subheader("1. Escopo Básico (Anotações de Reunião)")
                novo_cod = st.text_input("Título / Código da Revisão:", value=rev["codigo"], key=f"cod_{rev['id']}")
                novo_stat = st.selectbox("Status da Revisão:", STATUS, index=STATUS.index(rev["status"]) if rev["status"] in STATUS else 0,
                                         key=f"stat_{rev['id']}")
                escopo_raw = st.text_area("Pontos discutidos na reunião:", value=rev.get("escopo_basico", ""), height=260, key=f"raw_{rev['id']}")
                b1, b2 = st.columns(2)
                if b1.button("💾 Salvar Anotações", key=f"save_raw_{rev['id']}", use_container_width=True):
                    rev.update({"codigo": novo_cod, "status": novo_stat, "escopo_basico": escopo_raw})
                    salvar_dados()
                    st.success("Anotações salvas!")
                if b2.button("✨ Melhorar Escopo com Gemini", key=f"ai_rev_{rev['id']}", use_container_width=True):
                    rev.update({"codigo": novo_cod, "status": novo_stat, "escopo_basico": escopo_raw})
                    with st.spinner("Gemini transformando anotações em escopo técnico..."):
                        prompt_rev = f"""
                        Você é um Arquiteto Coordenador de Projetos. Pegue o escopo básico anotado durante uma reunião abaixo e gere uma versão profissional, técnica, clara e bem estruturada para compor um documento formal (.docx) de revisão de projeto.

                        Contexto do Projeto:
                        - Projeto: {proj['titulo']}
                        - Cliente: {proj['cliente']}
                        - Etapa: {proj['etapa_atual']}
                        - Revisão: {rev['codigo']}
                        - Briefing resumido: {proj.get('briefing', {}).get('programa_necessidades', '')}

                        Anotações básicas da reunião:
                        {escopo_raw}

                        Estruture o texto de forma limpa (sem usar asteriscos ** ou cerquilhas ### de markdown, use numeração e letras maiúsculas nos subtítulos) contendo:
                        1. OBJETIVO DA REVISÃO
                        2. DETALHAMENTO DAS ALTERAÇÕES E DIRETRIZES TÉCNICAS
                        3. PRANCHAS E ENTREGÁVEIS IMPACTADOS
                        4. RESPONSABILIDADES E PRÓXIMOS PASSOS
                        """
                        rev["escopo_aprimorado"] = chamar_gemini_ia(prompt_rev)
                        salvar_dados()
                        st.rerun()
            with col_r2:
                st.subheader("2. Versão Aprimorada & Exportação .docx")
                escopo_melhorado = st.text_area("Escopo profissional (edite antes de baixar):", value=rev.get("escopo_aprimorado", ""),
                                                height=335, key=f"enh_{rev['id']}")
                c_sv, c_dl = st.columns(2)
                if c_sv.button("💾 Salvar Versão Final", key=f"save_enh_{rev['id']}", use_container_width=True):
                    rev["escopo_aprimorado"] = escopo_melhorado
                    salvar_dados()
                    st.success("Versão final salva!")
                if HAS_DOCX:
                    rev["escopo_aprimorado"] = escopo_melhorado
                    c_dl.download_button("📥 Baixar em .docx", data=gerar_docx_revisao(proj, rev),
                                         file_name=f"{proj['titulo'].replace(' ', '_')}_{rev['codigo'].replace(' ', '_')}.docx",
                                         mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                         use_container_width=True)
                else:
                    c_dl.warning("Instale `python-docx` para baixar em .docx")
            with st.expander("🗑️ Excluir esta Revisão"):
                if st.button("Remover Revisão", key=f"del_rev_{rev['id']}"):
                    proj["revisoes"] = [r for r in revisoes if r["id"] != rev["id"]]
                    salvar_dados()
                    st.rerun()
        else:
            st.info("Nenhuma revisão cadastrada para este projeto ainda. Use o formulário acima para registrar a R00.")

# ==============================================================================
# TELA D: GRÁFICO DE ANDAMENTO & ETAPAS
# ==============================================================================
elif menu_atual == "📊 Gráfico de Andamento & Etapas":
    st.title("Andamento & Etapas")
    proj = obter_projeto_ativo()
    if proj:
        st.subheader("1. Esteira de Evolução das Etapas")
        etapas = db["etapas_projeto"]
        idx_atual = etapas.index(proj["etapa_atual"]) if proj["etapa_atual"] in etapas else 0
        pills = ""
        for idx, et in enumerate(etapas):
            if idx < idx_atual:
                pills += f'<span class="stage-pill stage-pill-done">✓ {esc(et)}</span>'
            elif idx == idx_atual:
                pills += f'<span class="stage-pill stage-pill-active">▶ {esc(et)} (Atual)</span>'
            else:
                pills += f'<span class="stage-pill">{esc(et)}</span>'
        st.markdown(pills, unsafe_allow_html=True)
        pct_etapa = int((idx_atual + 1) / len(etapas) * 100)
        st.progress(pct_etapa / 100, text=f"Avanço na linha de produção: {pct_etapa}% (etapa {idx_atual + 1} de {len(etapas)})")

        st.markdown("---")
        st.subheader("2. Cronograma de Etapas deste Projeto (Gantt)")
        with st.expander("➕ Adicionar Etapa ao Cronograma"):
            with st.form("form_etapa_gantt", clear_on_submit=True):
                gc1, gc2, gc3 = st.columns(3)
                g_etapa = gc1.selectbox("Etapa", db["etapas_projeto"])
                g_ini = gc2.date_input("Início da Etapa", date.today())
                g_fim = gc3.date_input("Fim da Etapa", date.today() + timedelta(days=7))
                if st.form_submit_button("Adicionar ao Cronograma"):
                    proj.setdefault("cronograma_etapas", []).append({"etapa": g_etapa, "inicio": g_ini.isoformat(), "fim": g_fim.isoformat()})
                    salvar_dados()
                    st.rerun()
        if proj.get("cronograma_etapas"):
            fig = px.timeline(pd.DataFrame(proj["cronograma_etapas"]), x_start="inicio", x_end="fim", y="etapa", color="etapa",
                              title=f"Cronograma — {proj['titulo']}", color_discrete_sequence=PALETA)
            fig.update_yaxes(autorange="reversed")
            fig.update_layout(showlegend=False)
            st.plotly_chart(estilo_plotly(fig, 340), use_container_width=True)

        st.markdown("---")
        st.subheader("3. Panorama Geral de Todos os Projetos")
        if db["projetos"]:
            fig_g = px.timeline(pd.DataFrame(db["projetos"]), x_start="data_inicio", x_end="data_fim", y="titulo", color="etapa_atual",
                                hover_data=["cliente", "categoria", "orcamento"], title="Linha do Tempo Global dos Projetos",
                                color_discrete_sequence=PALETA)
            fig_g.update_yaxes(autorange="reversed")
            st.plotly_chart(estilo_plotly(fig_g, 360), use_container_width=True)

# ==============================================================================
# TELA E: CATEGORIAS E ETAPAS DE PROJETOS
# ==============================================================================
elif menu_atual == "🏷️ Categorias & Etapas de Projetos":
    st.title("Categorias e Etapas")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.subheader("Categorias de Projetos (Tipologias)")
        nova_cat_p = st.text_input("Nova categoria (ex: Urbanismo, Reforma, Luminotécnico):")
        if st.button("➕ Adicionar Categoria de Projeto") and nova_cat_p.strip() and nova_cat_p.strip() not in db["categorias_projeto"]:
            db["categorias_projeto"].append(nova_cat_p.strip())
            salvar_dados()
            st.rerun()
        for i, cat in enumerate(db["categorias_projeto"]):
            cc1, cc2 = st.columns([0.85, 0.15])
            cc1.markdown(f"<div class='arqui-card' style='padding:10px;'>{esc(cat)}</div>", unsafe_allow_html=True)
            if cc2.button("🗑️", key=f"del_cat_p_{i}") and len(db["categorias_projeto"]) > 1:
                db["categorias_projeto"].pop(i)
                salvar_dados()
                st.rerun()
    with col_c2:
        st.subheader("Etapas do Fluxo de Trabalho")
        nova_etapa_p = st.text_input("Nova etapa (ex: Aprovação Prefeitura, Renderização):")
        if st.button("➕ Adicionar Etapa") and nova_etapa_p.strip() and nova_etapa_p.strip() not in db["etapas_projeto"]:
            db["etapas_projeto"].append(nova_etapa_p.strip())
            salvar_dados()
            st.rerun()
        for i, etp in enumerate(db["etapas_projeto"]):
            ce1, ce2, ce3 = st.columns([0.7, 0.15, 0.15])
            ce1.markdown(f"<div class='arqui-card' style='padding:10px;'>{i + 1}. {esc(etp)}</div>", unsafe_allow_html=True)
            if ce2.button("↑", key=f"up_etp_{i}") and i > 0:
                db["etapas_projeto"][i], db["etapas_projeto"][i - 1] = db["etapas_projeto"][i - 1], db["etapas_projeto"][i]
                salvar_dados()
                st.rerun()
            if ce3.button("🗑️", key=f"del_etp_{i}") and len(db["etapas_projeto"]) > 1:
                db["etapas_projeto"].pop(i)
                salvar_dados()
                st.rerun()

# ==============================================================================
# TELA F: CRIAÇÕES
# ==============================================================================
elif menu_atual == "➕ Nova Categoria de Criação":
    st.title("Nova Categoria em Criações")
    st.write("Crie categorias para organizar estudos, plugins, produtos ou ideias.")
    nova_cat_c = st.text_input("Nome da nova categoria (ex: Plugins AutoCAD, Renderizações, Ideias de Negócio):")
    if st.button("Criar Categoria") and nova_cat_c.strip() and nova_cat_c.strip() not in db["categorias_criacoes"]:
        db["categorias_criacoes"].append(nova_cat_c.strip())
        salvar_dados()
        st.session_state.menu_ativo = f"✨ {nova_cat_c.strip()}"
        st.rerun()
    st.markdown("---")
    st.subheader("Categorias existentes")
    for idx_cc, cat_c in enumerate(db["categorias_criacoes"]):
        c1, c2 = st.columns([0.85, 0.15])
        c1.markdown(f"<div class='arqui-card' style='padding:10px;'>✨ {esc(cat_c)}</div>", unsafe_allow_html=True)
        if cat_c not in ["Atuais", "Futuras", "Cursos em Andamento"] and c2.button("🗑️", key=f"del_cat_c_{idx_cc}"):
            db["categorias_criacoes"].pop(idx_cc)
            salvar_dados()
            st.rerun()

elif menu_atual.startswith("✨ "):
    categoria_atual = menu_atual.replace("✨ ", "", 1)
    st.title(f"Criações — {categoria_atual}")
    with st.expander(f"➕ Adicionar item em '{categoria_atual}'"):
        with st.form("form_nova_criacao", clear_on_submit=True):
            cr1, cr2 = st.columns([2, 1])
            c_titulo = cr1.text_input("Título da Criação / Curso / Ideia *")
            c_prog = cr2.slider("Progresso Atual (%)", 0, 100, 10)
            c_link = st.text_input("Link de Referência (Drive, GitHub, Udemy, etc.)")
            c_prox = st.text_input("Próximo Passo")
            c_notas = st.text_area("Anotações e Detalhes")
            if st.form_submit_button("Salvar Criação") and c_titulo.strip():
                db["criacoes"].append({"id": f"cria_{uuid.uuid4().hex[:6]}", "titulo": c_titulo.strip(), "categoria": categoria_atual,
                                       "progresso": c_prog, "link": c_link, "proximo_passo": c_prox, "anotacoes": c_notas})
                salvar_dados()
                st.rerun()

    itens_cat = [i for i in db["criacoes"] if i["categoria"] == categoria_atual]
    if itens_cat:
        for item in itens_cat:
            st.markdown(f"""<div class="arqui-card">
                <div style="display:flex;justify-content:space-between;align-items:center;">
                    <h4 style="margin:0;">{esc(item['titulo'])}</h4><span class="badge-tag">{int(item['progresso'])}% concluído</span></div>
                <p style="font-size:13px;color:#c9d4b8;margin:6px 0 4px 0;">Próximo passo: {esc(item.get('proximo_passo', '-'))}</p>
                <p style="font-size:13px;color:#7f8a6a;margin:0;">{esc(item.get('anotacoes', ''))}</p></div>""", unsafe_allow_html=True)
            st.progress(int(item["progresso"]) / 100)
            with st.expander(f"✏️ Editar '{item['titulo']}'"):
                ed_t = st.text_input("Título:", value=item["titulo"], key=f"ed_ct_{item['id']}")
                ed_cat = st.selectbox("Mover para categoria:", db["categorias_criacoes"],
                                      index=db["categorias_criacoes"].index(item["categoria"]) if item["categoria"] in db["categorias_criacoes"] else 0,
                                      key=f"ed_ccat_{item['id']}")
                ed_p = st.slider("Progresso (%):", 0, 100, int(item["progresso"]), key=f"ed_cp_{item['id']}")
                ed_l = st.text_input("Link:", value=item.get("link", ""), key=f"ed_cl_{item['id']}")
                ed_pr = st.text_input("Próximo passo:", value=item.get("proximo_passo", ""), key=f"ed_cpr_{item['id']}")
                ed_n = st.text_area("Anotações:", value=item.get("anotacoes", ""), key=f"ed_cn_{item['id']}")
                bc1, bc2, bc3 = st.columns(3)
                if bc1.button("💾 Salvar Edição", key=f"sv_c_{item['id']}", use_container_width=True):
                    item.update({"titulo": ed_t, "categoria": ed_cat, "progresso": ed_p, "link": ed_l, "proximo_passo": ed_pr, "anotacoes": ed_n})
                    salvar_dados()
                    st.rerun()
                if item.get("link"):
                    bc2.link_button("🔗 Abrir Link", item["link"], use_container_width=True)
                if bc3.button("🗑️ Excluir", key=f"dl_c_{item['id']}", use_container_width=True):
                    db["criacoes"] = [x for x in db["criacoes"] if x["id"] != item["id"]]
                    salvar_dados()
                    st.rerun()
    else:
        st.info(f"Nenhum item cadastrado em '{categoria_atual}' ainda.")

# ==============================================================================
# TELA G: AGENDA GOOGLE INTEGRADA
# ==============================================================================
elif menu_atual == "🗓️ Agenda Google & Categorias":
    st.title("Agenda Google & Categorias")
    col_ag1, col_ag2 = st.columns([2, 1])
    nomes_cats = [c["nome"] for c in db["categorias_agenda"]]

    with col_ag1:
        st.subheader("Novo Compromisso / Evento")
        with st.form("form_novo_evento_agenda", clear_on_submit=True):
            ac1, ac2 = st.columns([2, 1])
            ev_titulo = ac1.text_input("Título do Compromisso *")
            ev_cat = ac2.selectbox("Categoria da Agenda", nomes_cats)
            ac3, ac4 = st.columns(2)
            ev_data = ac3.date_input("Data", date.today())
            ev_hora = ac4.text_input("Horário (HH:MM)", value="09:00")
            ev_desc = st.text_area("Descrição / Local / Link da Reunião", height=80)
            if st.form_submit_button("➕ Adicionar à Agenda") and ev_titulo.strip():
                db["eventos_agenda"].append({"id": f"evt_{uuid.uuid4().hex[:6]}", "titulo": ev_titulo.strip(), "categoria": ev_cat,
                                             "data": ev_data.isoformat(), "hora": ev_hora.strip() or "09:00", "descricao": ev_desc})
                salvar_dados()
                st.rerun()
    with col_ag2:
        st.subheader("Categorias da Agenda")
        with st.form("form_nova_cat_agenda", clear_on_submit=True):
            nc_nome = st.text_input("Nova categoria (ex: Obra, Família, Estudos)")
            nc_cor = st.color_picker("Cor da categoria", value="#c9d4b8")
            if st.form_submit_button("Criar Categoria") and nc_nome.strip() and nc_nome.strip() not in nomes_cats:
                db["categorias_agenda"].append({"nome": nc_nome.strip(), "cor": nc_cor})
                salvar_dados()
                st.rerun()
        for c_ag in db["categorias_agenda"]:
            st.markdown(f"<span style='display:inline-block;width:12px;height:12px;border-radius:50%;background:{esc(c_ag['cor'])};margin-right:8px;'></span><b>{esc(c_ag['nome'])}</b>",
                        unsafe_allow_html=True)

    st.markdown("---")
    if ig.google_ativo():
        g1, g2 = st.columns(2)
        if g1.button("🔄 Importar próximos 30 dias do Google Agenda", use_container_width=True):
            try:
                st.session_state.eventos_google = ig.agenda_listar(30)
            except Exception as e:
                st.error(f"Erro ao ler a agenda: {e}")
        if g2.button("⬆️ Enviar prazos dos projetos ao Google Agenda", use_container_width=True):
            try:
                for p in db["projetos"]:
                    ig.agenda_upsert(p["id"], f"📐 Entrega: {p['titulo']}", p["data_fim"], "18:00",
                                     f"Cliente: {p['cliente']} | Etapa: {p['etapa_atual']}")
                st.success("Prazos sincronizados (sem duplicar).")
            except Exception as e:
                st.error(f"Erro ao enviar: {e}")
        if st.session_state.get("eventos_google"):
            st.markdown('<div class="mono-label" style="margin:12px 0 6px;">Do Google Agenda (próximos 30 dias)</div>', unsafe_allow_html=True)
            for ev in st.session_state.eventos_google:
                st.markdown(f"🗓️ **{esc(ev['data'])} {esc(ev['hora'])}** — {esc(ev['titulo'])}")
    else:
        st.info("Configure os secrets do Google para sincronizar de verdade (veja COMO_INTEGRAR.md).")

    filtro_cat = st.multiselect("Filtrar por categorias:", options=nomes_cats, default=nomes_cats)
    mapa_cores = {c["nome"]: c["cor"] for c in db["categorias_agenda"]}
    cat_entregas = "Entregas" if "Entregas" in nomes_cats else "Trabalho"

    todos = [{"id": e["id"], "origem": "evento", "titulo": e["titulo"], "categoria": e["categoria"], "data": e["data"],
              "hora": e.get("hora", "09:00"), "descricao": e.get("descricao", ""), "cor": mapa_cores.get(e["categoria"], "#c9d4b8")}
             for e in db["eventos_agenda"] if e["categoria"] in filtro_cat]
    if cat_entregas in filtro_cat:
        todos += [{"id": p["id"], "origem": "projeto", "titulo": f"📐 Entrega Projeto: {p['titulo']} ({p['cliente']})",
                   "categoria": cat_entregas, "data": p["data_fim"], "hora": "18:00",
                   "descricao": f"Etapa: {p['etapa_atual']} | {p.get('anotacoes', '')}", "cor": mapa_cores.get(cat_entregas, "#d9b46a")}
                  for p in db["projetos"]]
    todos.sort(key=lambda x: (x["data"], x["hora"]))

    st.subheader("Compromissos & Prazos")
    for comp in todos:
        c_info, c_gcal, c_del = st.columns([0.68, 0.24, 0.08])
        c_info.markdown(f"""<div class="arqui-card" style="border-left:4px solid {esc(comp['cor'])};padding:12px;">
            <span style="background:{esc(comp['cor'])};color:#0f110d;padding:2px 8px;border-radius:6px;font:600 11px 'Geist Mono',monospace;">{esc(comp['categoria'])}</span>
            <span class="mono-label" style="margin-left:8px;">{esc(comp['data'])} às {esc(comp['hora'])}</span>
            <h4 style="margin:6px 0 2px 0;">{esc(comp['titulo'])}</h4>
            <p style="font-size:12px;color:#7f8a6a;margin:0;">{esc(comp['descricao'])}</p></div>""", unsafe_allow_html=True)
        c_gcal.write("")
        if ig.google_ativo():
            if c_gcal.button("📅 Enviar", key=f"gcal_{comp['id']}", use_container_width=True):
                try:
                    _, acao = ig.agenda_upsert(comp["id"], comp["titulo"], comp["data"], comp["hora"], comp["descricao"])
                    st.toast(f"Evento {acao} no Google Agenda")
                except Exception as e:
                    st.error(f"Erro: {e}")
        else:
            c_gcal.link_button("📅 Enviar p/ Google Agenda", gerar_link_google_agenda(comp["titulo"], comp["data"], comp["hora"], comp["descricao"]),
                               use_container_width=True)
        if comp["origem"] == "evento":
            c_del.write("")
            if c_del.button("✕", key=f"del_ev_{comp['id']}"):
                db["eventos_agenda"] = [e for e in db["eventos_agenda"] if e["id"] != comp["id"]]
                salvar_dados()
                st.rerun()

    st.markdown("---")
    st.subheader("🌐 Visualização Direta do seu Google Agenda")
    embed_url = db["config"].get("google_calendar_embed", "").strip()
    if embed_url.startswith("http"):
        components.iframe(embed_url, height=600, scrolling=True)
    else:
        st.info("Para ver o Google Agenda ao vivo aqui: Google Agenda > Configurações da agenda > 'Integrar agenda' > copie o "
                "'Código de incorporação' (src=...) e cole em **Configurações > URL Embed Google Agenda** na barra lateral.")
