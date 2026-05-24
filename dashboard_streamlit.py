# -*- coding: utf-8 -*-
"""
Dashboard Simplifique Representações — Streamlit v6
Novidades: cards mobile-first, planilha completa admin, export Excel, Resumo IA
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import io
import re
from datetime import datetime, date

PLANILHA_ID = '1zRgd4x0Th67MYywSVWrE4W8R_3dK1Y7k4em_8YxrBMo'
API_KEY     = st.secrets["GOOGLE_API_KEY"]
ABA         = 'TICKETS'
DIAS_ATRASO = 7
MOTIVO_PRIO = 'prorrogação de boleto'

USUARIOS = {
    'admin':              {'senha': 'Simpl#2026!',  'perfil': 'admin', 'rep': None},
    'juliana':            {'senha': 'juliana@2026', 'perfil': 'admin', 'rep': None},
    'vinicius_varrichio': {'senha': 'Vn#8xK2$mQ',  'perfil': 'rep',   'rep': 'VINICIUS VARRICHIO'},
    'thiago_duroes':      {'senha': 'Th@5wL9#pR',   'perfil': 'rep',   'rep': 'THIAGO DUROES'},
    'anderson_costa':     {'senha': 'An@6kM2#qZ',   'perfil': 'rep',   'rep': 'ANDERSON COSTA'},
    'fernando_azevedo':   {'senha': 'Fe#9rT4$wP',   'perfil': 'rep',   'rep': 'FERNANDO AZEVEDO'},
    'antonio_sidney':     {'senha': 'At$7yH3!nL',   'perfil': 'rep',   'rep': 'ANTONIO FERREIRA'},
    'heygla_silva':       {'senha': 'Hg@2xQ8#kV',   'perfil': 'rep',   'rep': 'HEYGLA SILVA'},
}

st.set_page_config(page_title='Simplifique Representações', page_icon='📋',
                   layout='wide', initial_sidebar_state='expanded')

st.markdown("""
<style>
div[data-testid="metric-container"] {
    background:#ffffff; border:1px solid #e5e7eb;
    border-radius:8px; padding:12px 16px;
}
div[data-testid="metric-container"] label {
    font-size:11px !important; color:#6b7a99 !important;
    text-transform:uppercase; letter-spacing:0.4px;
}
.block-container { padding-top:1.5rem; padding-bottom:1rem; }
.card { background:#fff; border:1px solid #e5e7eb; border-radius:12px; padding:14px 16px; margin-bottom:12px; }
.card-prio   { border-left:4px solid #dc2626; background:#fff8f8; }
.card-atraso { border-left:4px solid #f59e0b; background:#fffbeb; }
.card-normal { border-left:4px solid #3b82f6; }
.card-titulo { font-size:14px; font-weight:700; color:#1a2540; margin-bottom:4px; }
.card-meta   { font-size:12px; color:#6b7a99; margin-bottom:8px; }
.card-resumo { font-size:13px; color:#374151; line-height:1.5; }
.badge { display:inline-block; border-radius:6px; padding:2px 8px; font-size:11px; font-weight:600; margin-right:4px; }
.badge-red   { background:#fee2e2; color:#dc2626; }
.badge-amber { background:#fef3c7; color:#d97706; }
.badge-blue  { background:#dbeafe; color:#1d4ed8; }
.badge-gray  { background:#f3f4f6; color:#6b7280; }
</style>
""", unsafe_allow_html=True)

# ─── FUNÇÕES AUXILIARES ───────────────────────────────────────────────────────

def limpar_nome(nome):
    if not nome: return ''
    return re.sub(r'^R\.?E\.?\s*[-–]\s*', '', nome, flags=re.IGNORECASE).strip()

def limpar_gerente(nome):
    if not nome: return ''
    return re.sub(r'^G\.?E\.?\s*[-–]\s*', '', nome, flags=re.IGNORECASE).strip()

def is_resolvido(status):
    return 'resolvido' in str(status).lower()

def is_prioridade(motivo):
    return MOTIVO_PRIO in str(motivo).lower()

def deve_excluir_prioridade(row):
    if row['Resolvido']: return True
    if str(row['PendenteCom']).strip().lower() == 'finalizado': return True
    return False

def parsear_data(s):
    if not s or str(s).strip() == '': return None
    try:
        p = str(s).strip().split('/')
        if len(p) == 3:
            return date(int(p[2]), int(p[1]), int(p[0]))
    except: pass
    return None

def is_pendente_rep_cliente(row):
    status_dpto = ('dpto. representante' in str(row['Status']).lower() or
                   'dpto representante' in str(row['Status']).lower())
    pend_cli_rep = str(row['PendenteCom']).strip().lower() in ['cliente', 'representante']
    return status_dpto or pend_cli_rep

def bar_chart_pct(df_plot, x_col, y_col, color, height=280):
    total = df_plot[x_col].sum() or 1
    df_plot = df_plot.copy()
    df_plot['pct']   = (df_plot[x_col] / total * 100).round(1)
    df_plot['label'] = df_plot.apply(lambda r: f"{int(r[x_col])} ({r['pct']}%)", axis=1)
    fig = px.bar(df_plot, x=x_col, y=y_col, orientation='h',
                 text='label', color_discrete_sequence=[color])
    fig.update_traces(textposition='outside', cliponaxis=False)
    fig.update_layout(margin=dict(t=10,b=10,l=10,r=90), height=height,
                      yaxis_title='', xaxis_title='', showlegend=False)
    return fig

def gerar_excel_tabela(df_in, nome_aba='Tickets', titulo=''):
    output = io.BytesIO()
    try:
        import xlsxwriter
        wb  = xlsxwriter.Workbook(output, {'in_memory': True})
        fmt_tit  = wb.add_format({'bold':True,'font_size':12,'font_color':'#FFFFFF','bg_color':'#C0392B','align':'left','valign':'vcenter'})
        fmt_hdr  = wb.add_format({'bold':True,'font_color':'#FFFFFF','bg_color':'#1a2540','align':'center','valign':'vcenter','border':1,'font_size':10,'text_wrap':True})
        fmt_data = wb.add_format({'border':1,'valign':'top','text_wrap':True,'font_size':10})
        fmt_alt  = wb.add_format({'border':1,'valign':'top','text_wrap':True,'font_size':10,'bg_color':'#F0F4FF'})
        fmt_prio = wb.add_format({'border':1,'valign':'top','text_wrap':True,'font_size':10,'bg_color':'#FECACA','font_color':'#991B1B'})
        ws = wb.add_worksheet(nome_aba[:31])
        hoje = datetime.now().strftime('%d/%m/%Y %H:%M')
        titulo_final = titulo or f'Exportação — {hoje}'
        ncols = len(df_in.columns)
        ws.set_row(0, 22); ws.set_row(1, 30)
        ws.merge_range(0, 0, 0, ncols-1, titulo_final, fmt_tit)
        for ci, col in enumerate(df_in.columns):
            ws.write(1, ci, col, fmt_hdr)
            larg = 14
            col_l = col.lower()
            if 'título' in col_l or 'resumo' in col_l or 'observ' in col_l: larg = 50
            elif 'ticket' in col_l: larg = 10
            elif 'representante' in col_l or 'gerente' in col_l: larg = 22
            elif 'status' in col_l or 'motivo' in col_l or 'pendente' in col_l: larg = 20
            elif 'link' in col_l or 'url' in col_l: larg = 40
            ws.set_column(ci, ci, larg)
        for ri, (_, row) in enumerate(df_in.iterrows()):
            is_pri = '🚨' in str(row.get('🚨',''))
            fmt_row = fmt_prio if is_pri else (fmt_alt if ri % 2 == 0 else fmt_data)
            for ci, col in enumerate(df_in.columns):
                val = row[col]
                ws.write(ri+2, ci, str(val) if val is not None else '', fmt_row)
        wb.close()
    except ImportError:
        output.write(df_in.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig'))
    output.seek(0)
    return output

def gerar_excel_report(df_cli, df_pos, rep_nome):
    output = io.BytesIO()
    try:
        import xlsxwriter
        workbook  = xlsxwriter.Workbook(output, {'in_memory': True})
        fmt_tit   = workbook.add_format({'bold':True,'font_size':13,'font_color':'#FFFFFF','bg_color':'#C0392B','align':'left','valign':'vcenter'})
        fmt_hdr   = workbook.add_format({'bold':True,'font_color':'#FFFFFF','bg_color':'#C0392B','align':'center','valign':'vcenter','border':1,'font_size':10})
        fmt_data  = workbook.add_format({'border':1,'valign':'top','text_wrap':True,'font_size':10})
        fmt_alt   = workbook.add_format({'border':1,'valign':'top','text_wrap':True,'font_size':10,'bg_color':'#FFF3CD'})
        fmt_prio  = workbook.add_format({'border':1,'valign':'top','text_wrap':True,'font_size':10,'bg_color':'#FECACA','font_color':'#991B1B'})
        hoje = datetime.now().strftime('%d/%m/%Y')
        col_w = {'Nº Ticket':12,'Título':55,'Data Abertura':14,'Dias em Aberto':14,'Pendente Com':18,'Observação':45,'Motivo':22,'Status':22}
        secoes = [
            (df_cli, ['Ticket','Titulo','DataAbertura','DiasAberto','PendenteCom','Observacao'],
                     ['Nº Ticket','Título','Data Abertura','Dias em Aberto','Pendente Com','Observação'], 'Pendentes - Cliente'),
            (df_pos, ['Ticket','Titulo','DataAbertura','DiasAberto','PendenteCom','Motivo'],
                     ['Nº Ticket','Título','Data Abertura','Dias em Aberto','Pendente Com','Motivo'],    'Pendentes - Pos-Venda Comercial'),
        ]
        for df_sec, cols, labels, sheet in secoes:
            ws = workbook.add_worksheet(sheet)
            ws.set_row(0, 24); ws.set_row(1, 20)
            ws.merge_range(0, 0, 0, len(labels)-1, f'REPORT SEMANAL — PÓS-VENDAS PENDENTES REP. {rep_nome.upper()} — {hoje}', fmt_tit)
            for ci, lbl in enumerate(labels):
                ws.write(1, ci, lbl, fmt_hdr)
                ws.set_column(ci, ci, col_w.get(lbl, 18))
            df_exp = df_sec[[c for c in cols if c in df_sec.columns]].copy()
            for ri, (_, row) in enumerate(df_exp.iterrows()):
                is_pri  = is_prioridade(str(row.get('Motivo', '')))
                fmt_row = fmt_prio if is_pri else (fmt_alt if ri % 2 == 0 else fmt_data)
                for ci, col in enumerate(cols):
                    val = row.get(col, '') if col in df_sec.columns else ''
                    ws.write(ri + 2, ci, str(val) if val else '', fmt_row)
        workbook.close()
    except ImportError:
        output.write(df_cli.to_csv(index=False).encode('utf-8-sig'))
    output.seek(0)
    return output

# ─── CARGA DE DADOS ───────────────────────────────────────────────────────────

@st.cache_data(ttl=120)
def verificar_ultima_atualizacao():
    url = f'https://sheets.googleapis.com/v4/spreadsheets/{PLANILHA_ID}/values/{ABA}!U:U?key={API_KEY}'
    try:
        resp   = requests.get(url, timeout=10)
        valores = resp.json().get('values', [])
        vals   = [v[0] for v in valores[1:] if v and v[0].strip()]
        return vals[-1] if vals else ''
    except:
        return ''

@st.cache_data(ttl=86400)
def carregar_dados():
    url = f'https://sheets.googleapis.com/v4/spreadsheets/{PLANILHA_ID}/values/{ABA}!A:V?key={API_KEY}'
    try:
        resp = requests.get(url, timeout=15)
        data = resp.json()
        if 'error' in data:
            return None, f"Erro: {data['error']['message']}"
        valores = data.get('values', [])
        if len(valores) < 2: return pd.DataFrame(), None
        cab = ['Representante','Ticket','Titulo','DataAbertura','DataAtualizacao',
               'HoraAtualizacao','DiasAberto','Status','PendenteCom','Motivo',
               'Observacao','PrimeiroEmail','UltimoHistorico','QuemRespondeu',
               'CodCliente','RazaoSocial','NF','Pedido','Gerente','URL',
               'UltimaAtualizacao','ResumoIA']
        linhas = []
        for row in valores[1:]:
            while len(row) < 22: row.append('')
            linhas.append(row[:22])
        df = pd.DataFrame(linhas, columns=cab)
        df = df[df['Ticket'].str.strip() != '']
        df['Ticket']     = df['Ticket'].str.replace('#', '', regex=False).str.strip()
        df['DiasAberto'] = pd.to_numeric(df['DiasAberto'], errors='coerce').fillna(0).astype(int)
        df['RepNome']    = df['Representante'].apply(limpar_nome)
        df['GerNome']    = df['Gerente'].apply(limpar_gerente)
        df['Prioridade'] = df['Motivo'].apply(is_prioridade)
        df['Resolvido']  = df['Status'].apply(is_resolvido)
        df['Atrasado']   = (~df['Resolvido']) & (df['DiasAberto'] >= DIAS_ATRASO)
        df['DataAberturaDate'] = df['DataAbertura'].apply(parsear_data)
        df['PrioAtiva']  = df.apply(lambda r: r['Prioridade'] and not deve_excluir_prioridade(r), axis=1)
        df['PendRepCli'] = df.apply(is_pendente_rep_cliente, axis=1)
        df['ResumoIA']   = df['ResumoIA'].apply(
            lambda x: re.sub(r'^\[ATU:[^\]]+\]\s*', '', str(x)).strip() if x else ''
        )
        return df, None
    except Exception as e:
        return None, str(e)

# ─── TELA DE LOGIN ────────────────────────────────────────────────────────────
def tela_login():
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("""
<div style="background:#fff;border:1px solid #e5e7eb;border-radius:12px;
            padding:36px 40px;box-shadow:0 4px 24px rgba(0,0,0,0.08);margin-top:60px">
  <div style="text-align:center;margin-bottom:24px">
    <div style="font-size:36px">📋</div>
    <div style="font-size:18px;font-weight:700;color:#1a2540;margin-top:6px">Simplifique Representações</div>
    <div style="font-size:12px;color:#6b7a99;margin-top:4px">Pós-Vendas — Acesso restrito</div>
  </div>
</div>
""", unsafe_allow_html=True)
        with st.form('login_form'):
            usuario = st.text_input('👤 Usuário', placeholder='seu usuário').strip().lower()
            senha   = st.text_input('🔒 Senha',   placeholder='sua senha', type='password')
            entrar  = st.form_submit_button('Entrar', use_container_width=True)
            if entrar:
                if usuario in USUARIOS and USUARIOS[usuario]['senha'] == senha:
                    st.session_state.logado   = True
                    st.session_state.usuario  = usuario
                    st.session_state.perfil   = USUARIOS[usuario]['perfil']
                    st.session_state.rep_fixo = USUARIOS[usuario]['rep']
                    st.rerun()
                else:
                    st.error('Usuário ou senha incorretos.')

if 'logado' not in st.session_state:
    st.session_state.logado = False
if not st.session_state.logado:
    tela_login()
    st.stop()

if 'filtros' not in st.session_state:
    st.session_state.filtros = {
        'rep':'Todos','ger':'Todos','status':'Todos','pend':'Todos','mot':'Todos',
        'so_atraso':False,'so_prio':False,'dab1':None,'dab2':None
    }

perfil_atual  = st.session_state.get('perfil', 'rep')
usuario_atual = st.session_state.get('usuario', '')
rep_fixo      = st.session_state.get('rep_fixo', None)

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('### 📋 Simplifique Representações')
    icone = '👑' if perfil_atual == 'admin' else '👤'
    st.caption(f'{icone} **{usuario_atual.capitalize()}**')
    st.markdown('---')
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button('🔄 Atualizar', use_container_width=True):
            carregar_dados.clear()
            verificar_ultima_atualizacao.clear()
            st.rerun()
    with col_btn2:
        if st.button('🧹 Limpar', use_container_width=True):
            st.session_state.filtros = {
                'rep':'Todos','ger':'Todos','status':'Todos','pend':'Todos','mot':'Todos',
                'so_atraso':False,'so_prio':False,'dab1':None,'dab2':None
            }
            for key in ['sb_rep','sb_ger','sb_status','sb_pend','sb_mot','cb_atraso','cb_prio','dab1_in','dab2_in']:
                if key in st.session_state: del st.session_state[key]
            st.rerun()
    st.markdown('---')
    if st.button('🚪 Sair', use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
    st.markdown('### Filtros')

ult_check = verificar_ultima_atualizacao()
if 'ultima_atu_check' not in st.session_state:
    st.session_state.ultima_atu_check = ult_check
if ult_check and ult_check != st.session_state.ultima_atu_check:
    st.session_state.ultima_atu_check = ult_check
    carregar_dados.clear()

df_raw, erro = carregar_dados()
if erro: st.error(f'❌ {erro}'); st.stop()
if df_raw is None or df_raw.empty: st.warning('Nenhum dado.'); st.stop()

with st.sidebar:
    def idx(lst, val): return lst.index(val) if val in lst else 0
    if perfil_atual == 'rep':
        st.info(f'📊 Visualizando: **{rep_fixo}**')
        rep_sel = 'Todos'; ger_sel = 'Todos'; status_sel = 'Todos'
        pend_sel = 'Todos'; mot_sel = 'Todos'
        dab1 = None; dab2 = None; so_atraso = False; so_prio = False
    else:
        reps_l   = ['Todos'] + sorted(df_raw['RepNome'].dropna().unique().tolist())
        gers_l   = ['Todos'] + sorted(df_raw['GerNome'].dropna().unique().tolist())
        status_l = ['Todos','Excluir Resolvidos'] + sorted(df_raw['Status'].dropna().unique().tolist())
        pend_l   = ['Todos'] + sorted(df_raw['PendenteCom'].replace('',pd.NA).dropna().unique().tolist())
        mot_l    = ['Todos'] + sorted(df_raw['Motivo'].replace('',pd.NA).dropna().unique().tolist())
        rep_sel    = st.selectbox('Representante',    reps_l,   index=idx(reps_l,   st.session_state.filtros['rep']),    key='sb_rep')
        ger_sel    = st.selectbox('Gerente Regional', gers_l,   index=idx(gers_l,   st.session_state.filtros['ger']),    key='sb_ger')
        status_sel = st.selectbox('Status',           status_l, index=idx(status_l, st.session_state.filtros['status']), key='sb_status')
        pend_sel   = st.selectbox('Pendente Com',     pend_l,   index=idx(pend_l,   st.session_state.filtros['pend']),   key='sb_pend')
        mot_sel    = st.selectbox('Motivo',           mot_l,    index=idx(mot_l,    st.session_state.filtros['mot']),    key='sb_mot')
        st.markdown('**Período de abertura**')
        dab1 = st.date_input('De',  value=st.session_state.filtros.get('dab1'), format='DD/MM/YYYY', key='dab1_in')
        dab2 = st.date_input('Até', value=st.session_state.filtros.get('dab2'), format='DD/MM/YYYY', key='dab2_in')
        st.markdown('---')
        so_atraso = st.checkbox('Somente em atraso (+7d)', value=st.session_state.filtros['so_atraso'], key='cb_atraso')
        so_prio   = st.checkbox('Somente prioridades',     value=st.session_state.filtros['so_prio'],   key='cb_prio')
        st.session_state.filtros = {
            'rep':rep_sel,'ger':ger_sel,'status':status_sel,'pend':pend_sel,'mot':mot_sel,
            'so_atraso':so_atraso,'so_prio':so_prio,'dab1':dab1,'dab2':dab2
        }
    st.markdown('---')
    st.caption(f'Total na planilha: **{len(df_raw)}** tickets')

df = df_raw.copy()
if rep_fixo:
    df = df[df['Representante'].str.upper().str.contains(rep_fixo.upper(), na=False)]
elif rep_sel != 'Todos':              df = df[df['RepNome'] == rep_sel]
if ger_sel    != 'Todos':              df = df[df['GerNome'] == ger_sel]
if status_sel == 'Excluir Resolvidos': df = df[~df['Resolvido']]
elif status_sel != 'Todos':            df = df[df['Status'] == status_sel]
if pend_sel   != 'Todos':              df = df[df['PendenteCom'] == pend_sel]
if mot_sel    != 'Todos':              df = df[df['Motivo'] == mot_sel]
if so_atraso:                          df = df[df['Atrasado']]
if so_prio:                            df = df[df['PrioAtiva']]
if dab1:
    df = df[df['DataAberturaDate'].apply(lambda d: d is not None and d >= dab1)]
if dab2:
    df = df[df['DataAberturaDate'].apply(lambda d: d is not None and d <= dab2)]

COLS_TABELA = ['Ticket','Titulo','RepNome','GerNome','DataAbertura','DiasAberto',
               'Status','Motivo','PendenteCom','Observacao','DataAtualizacao','URL','PrioAtiva']
LABELS_TAB  = ['Ticket','Título','Representante','Gerente','Abertura','Dias',
               'Status','Motivo','Pendente Com','Observação','Últ. Atualiz.','Link','PrioAtiva']

def preparar_tabela(df_in):
    df_t = df_in[COLS_TABELA].copy()
    df_t.columns = LABELS_TAB
    df_t['🚨'] = df_t['PrioAtiva'].apply(lambda x: '🚨' if x else '')
    df_t = df_t.drop(columns=['PrioAtiva'])
    return df_t[['🚨','Ticket','Título','Representante','Gerente','Abertura','Dias',
                 'Status','Motivo','Pendente Com','Observação','Últ. Atualiz.','Link']]

def mostrar_tabela(df_t, key_suffix='', titulo_excel=''):
    st.dataframe(df_t.sort_values('Dias', ascending=False),
                 use_container_width=True, hide_index=True,
                 column_config={
                     'Dias': st.column_config.NumberColumn('Dias', format='%d'),
                     'Link': st.column_config.LinkColumn('Link', display_text='Abrir'),
                     '🚨':   st.column_config.TextColumn('🚨', width='small'),
                 })
    excel_data = gerar_excel_tabela(df_t.sort_values('Dias', ascending=False),
                                    nome_aba='Tickets', titulo=titulo_excel or f'Exportação {datetime.now().strftime("%d/%m/%Y")}')
    st.download_button(
        '📥 Exportar Excel', data=excel_data,
        file_name=f'export_{datetime.now().strftime("%d-%m-%Y")}_{key_suffix}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        key=f'dl_{key_suffix}'
    )

col_tit, col_info = st.columns([3, 2])
with col_tit:
    st.markdown('## 📋 Simplifique Representações')
with col_info:
    st.markdown(' ')
    ult_atu = ''
    if not df_raw.empty and 'UltimaAtualizacao' in df_raw.columns:
        vals = df_raw['UltimaAtualizacao'].replace('', pd.NA).dropna()
        if not vals.empty: ult_atu = vals.iloc[-1]
    st.caption(
        f'📅 Dados resolvidos a partir de **08/05/2026** &nbsp;&nbsp;|&nbsp;&nbsp; '
        f'🔄 Última atualização: **{ult_atu or datetime.now().strftime("%d/%m/%Y %H:%M")}**'
    )

if perfil_atual == 'admin':
    aba1, aba2, aba3, aba4, aba5, aba6 = st.tabs([
        '📊 Visão Geral', '🔍 Análise Detalhada',
        '📱 Meus Tickets', '🚨 Prioridades',
        '📧 Report Semanal', '🗂️ Planilha Completa',
    ])
else:
    aba1, aba2, aba3, aba4 = st.tabs([
        '📊 Visão Geral', '🔍 Análise Detalhada',
        '📱 Meus Tickets', '🚨 Prioridades',
    ])
    aba5 = None; aba6 = None

# ════════════ ABA 1 — VISÃO GERAL ════════════
with aba1:
    mask_resolvido = (
        df['Status'].str.strip().str.lower().str.contains('resolvido', na=False) |
        (df['PendenteCom'].str.strip().str.lower() == 'finalizado')
    )
    mask_aberto   = ~mask_resolvido
    mask_dpto_rep = (
        df['Status'].str.strip().str.lower().str.contains('dpto. representante', na=False) |
        df['Status'].str.strip().str.lower().str.contains('dpto representante', na=False)
    )
    mask_pend_cli = mask_dpto_rep | (df['PendenteCom'].str.strip().str.lower() == 'cliente')
    total       = len(df)
    abertos     = int(mask_aberto.sum())
    resolvidos  = int(mask_resolvido.sum())
    atrasados   = int((mask_aberto & (df['DiasAberto'] >= DIAS_ATRASO)).sum())
    prio_ativas = int((mask_aberto & df['Motivo'].apply(is_prioridade)).sum())
    pend_cli    = int((mask_aberto & mask_pend_cli).sum())
    pend_pos    = int((mask_aberto & ~mask_pend_cli).sum())
    c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
    c1.metric('📋 Total',           total)
    c2.metric('📂 Em Aberto',       abertos,     help='Excluídos: Resolvido e Finalizado')
    c3.metric('✅ Resolvidos',      resolvidos,  help='Status Resolvido e/ou Pendente Com: Finalizado')
    c4.metric('⏰ Atraso +7d',      atrasados,   delta=f'-{atrasados}'   if atrasados   else None, delta_color='inverse')
    c5.metric('🚨 Prioridades',     prio_ativas, delta=f'-{prio_ativas}' if prio_ativas else None, delta_color='inverse')
    c6.metric('👤 Pend. Cliente',   pend_cli,    delta=f'-{pend_cli}'    if pend_cli    else None, delta_color='inverse')
    c7.metric('🏢 Pend. Pós-Venda', pend_pos)
    st.markdown('---')
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown('### Tickets por Status')
        df_st = df['Status'].value_counts().reset_index(); df_st.columns = ['Status','Qtd']
        fig_st = px.pie(df_st, names='Status', values='Qtd', color_discrete_sequence=px.colors.qualitative.Set2, hole=0.4)
        fig_st.update_traces(textinfo='label+percent', textposition='inside')
        fig_st.update_layout(showlegend=True, margin=dict(t=10,b=10,l=10,r=10), height=270)
        st.plotly_chart(fig_st, use_container_width=True)
    with col_b:
        st.markdown('### Pendente Com')
        df_pend = df[df['PendenteCom']!='']['PendenteCom'].value_counts().reset_index(); df_pend.columns = ['Pendente Com','Qtd']
        if not df_pend.empty:
            cores = ['#ef4444' if p=='Cliente' else '#f59e0b' for p in df_pend['Pendente Com']]
            fig_pend = bar_chart_pct(df_pend, 'Qtd', 'Pendente Com', '#f59e0b', height=270)
            fig_pend.update_traces(marker_color=cores)
            st.plotly_chart(fig_pend, use_container_width=True)
    col_c, col_d = st.columns(2)
    with col_c:
        st.markdown('### Abertos por Representante')
        df_rep = df[~df['Resolvido']].groupby('RepNome').size().reset_index(name='Abertos').sort_values('Abertos', ascending=True)
        if not df_rep.empty:
            st.plotly_chart(bar_chart_pct(df_rep,'Abertos','RepNome','#1a56db',height=320), use_container_width=True)
    with col_d:
        st.markdown('### Atraso por Representante')
        df_at = df[df['Atrasado']].groupby('RepNome').size().reset_index(name='Atrasados').sort_values('Atrasados', ascending=True)
        if df_at.empty: st.info('Nenhum ticket em atraso.')
        else: st.plotly_chart(bar_chart_pct(df_at,'Atrasados','RepNome','#ef4444',height=320), use_container_width=True)
    st.markdown('### KPIs por Representante')
    kpi = df.groupby('RepNome').agg(
        Total=('Ticket','count'), Abertos=('Resolvido', lambda x: (~x).sum()),
        Atraso=('Atrasado','sum'), Prioridades=('PrioAtiva','sum'),
        PendCliente=('PendenteCom', lambda x: (x=='Cliente').sum()),
        PendPosVenda=('PendenteCom', lambda x: (x=='Pós-venda/Comercial').sum()),
    ).reset_index().rename(columns={'RepNome':'Representante','PendCliente':'Pend. Cliente','PendPosVenda':'Pend. Pós-Venda'})
    st.dataframe(kpi.sort_values('Abertos', ascending=False), use_container_width=True, hide_index=True)

# ════════════ ABA 2 — ANÁLISE DETALHADA ════════════
with aba2:
    col_f1, col_f2 = st.columns([3,1])
    with col_f1:
        busca = st.text_input('🔍 Buscar por ticket, título, NF, cliente ou observação', placeholder='Ex: 10482 ou boleto ou Drogapop...')
    with col_f2:
        ordenar = st.selectbox('Ordenar por', ['Dias (maior)','Dias (menor)','Ticket','Representante'])
    df_det = df.copy()
    if busca:
        b = busca.lower()
        df_det = df_det[
            df_det['Ticket'].str.lower().str.contains(b, na=False) |
            df_det['Titulo'].str.lower().str.contains(b, na=False) |
            df_det['Observacao'].str.lower().str.contains(b, na=False) |
            df_det['NF'].str.lower().str.contains(b, na=False) |
            df_det['RazaoSocial'].str.lower().str.contains(b, na=False)
        ]
    ord_map = {'Dias (maior)':('DiasAberto',False),'Dias (menor)':('DiasAberto',True),'Ticket':('Ticket',True),'Representante':('RepNome',True)}
    col_ord, asc_ord = ord_map[ordenar]
    df_det = df_det.sort_values(by=col_ord, ascending=asc_ord)
    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric('Filtrado', len(df_det)); m2.metric('Abertos', int((~df_det['Resolvido']).sum()))
    m3.metric('Resolvidos', int(df_det['Resolvido'].sum())); m4.metric('Atraso', int(df_det['Atrasado'].sum()))
    m5.metric('Prioridades', int(df_det['PrioAtiva'].sum()))
    st.markdown('---')
    with st.expander(f'📋 Todos os tickets ({len(df_det)})', expanded=True):
        mostrar_tabela(preparar_tabela(df_det), 'todos', 'Todos os Tickets')
    df_det_prio = df_det[df_det['PrioAtiva']]
    with st.expander(f'🚨 Prioridades ({len(df_det_prio)})', expanded=True):
        if df_det_prio.empty: st.success('Nenhuma prioridade ativa.')
        else: mostrar_tabela(preparar_tabela(df_det_prio), 'prio', 'Prioridades')
    df_det_rc = df_det[df_det['PendRepCli'] & ~df_det['Resolvido']]
    with st.expander(f'👤 Pendentes com Representante / Cliente ({len(df_det_rc)})', expanded=True):
        st.caption('Inclui: Status = Dpto. Representante  OU  Pendente Com = Cliente ou Representante')
        if df_det_rc.empty: st.success('Nenhum ticket nesta condição.')
        else: mostrar_tabela(preparar_tabela(df_det_rc), 'rep_cli', 'Pendentes Rep/Cliente')
    df_det_res = df_det[
        df_det['Status'].str.strip().str.lower().str.contains('resolvido', na=False) |
        (df_det['PendenteCom'].str.strip().str.lower() == 'finalizado')
    ]
    with st.expander(f'✅ Resolvidos ({len(df_det_res)})', expanded=False):
        if df_det_res.empty: st.success('Nenhum ticket resolvido.')
        else: mostrar_tabela(preparar_tabela(df_det_res), 'resolvidos', 'Resolvidos')

# ════════════ ABA 3 — MEUS TICKETS (cards mobile-first) ════════════
with aba3:
    if perfil_atual == 'rep':
        rep_meus_nome = df_raw[df_raw['Representante'].str.upper().str.contains(
            rep_fixo.upper(), na=False)]['RepNome'].iloc[0] if rep_fixo and not df_raw[
            df_raw['Representante'].str.upper().str.contains(rep_fixo.upper(), na=False)].empty else rep_fixo
        st.markdown(f'### 📱 Meus Tickets — {rep_meus_nome}')
        df_meus = df_raw[df_raw['Representante'].str.upper().str.contains(rep_fixo.upper(), na=False)].copy()
    else:
        reps_disp3 = sorted(df_raw['RepNome'].dropna().unique().tolist())
        rep_meus_nome = st.selectbox('👤 Representante', reps_disp3, key='rep_meus')
        df_meus = df_raw[df_raw['RepNome'] == rep_meus_nome].copy()

    df_ab = df_meus[~df_meus['Resolvido']]
    k1, k2, k3, k4 = st.columns(4)
    k1.metric('📂 Abertos', len(df_ab))
    k2.metric('⏰ Atraso', int(df_ab['Atrasado'].sum()))
    k3.metric('🚨 Prioridades', int(df_ab['PrioAtiva'].sum()))
    k4.metric('👤 Pend. Cliente', int((df_ab['PendenteCom']=='Cliente').sum()))
    st.markdown('---')

    busca3 = st.text_input('🔍 Buscar por NF, cliente, ticket ou palavra-chave',
                            placeholder='Ex: 1009005 ou Drogapop ou boleto...', key='busca_meus')
    df_vis = df_ab.copy()
    if busca3:
        b3 = busca3.lower()
        df_vis = df_vis[
            df_vis['Ticket'].str.lower().str.contains(b3, na=False) |
            df_vis['Titulo'].str.lower().str.contains(b3, na=False) |
            df_vis['NF'].str.lower().str.contains(b3, na=False) |
            df_vis['RazaoSocial'].str.lower().str.contains(b3, na=False) |
            df_vis['ResumoIA'].str.lower().str.contains(b3, na=False) |
            df_vis['Observacao'].str.lower().str.contains(b3, na=False)
        ]
    df_vis = df_vis.sort_values('DiasAberto', ascending=False)

    if df_vis.empty:
        st.success('✅ Nenhum ticket em aberto encontrado.')
    else:
        df_prio3   = df_vis[df_vis['PrioAtiva']]
        df_atraso3 = df_vis[~df_vis['PrioAtiva'] & df_vis['Atrasado']]
        df_normal3 = df_vis[~df_vis['PrioAtiva'] & ~df_vis['Atrasado']]

        def render_cards(df_c, classe):
            for _, row in df_c.iterrows():
                resumo = str(row.get('ResumoIA', '')).strip()
                titulo = str(row['Titulo'])[:80] + ('…' if len(str(row['Titulo'])) > 80 else '')
                dias   = int(row['DiasAberto'])
                pend   = str(row['PendenteCom']) if row['PendenteCom'] else '—'
                obs    = str(row['Observacao'])[:80] if row['Observacao'] else ''
                url    = str(row['URL'])
                badge_dias = (f'<span class="badge badge-red">⏰ {dias}d</span>'
                              if dias >= 7 else f'<span class="badge badge-gray">📅 {dias}d</span>')
                badge_pend = f'<span class="badge badge-amber">👤 {pend}</span>' if pend != '—' else ''
                badge_prio = '<span class="badge badge-red">🚨 Prioridade</span>' if row['PrioAtiva'] else ''
                resumo_html = (f'<div class="card-resumo">{resumo}</div>' if resumo
                               else (f'<div class="card-resumo" style="color:#9ca3af;font-style:italic">📝 {obs}</div>' if obs else ''))
                st.markdown(f'''
                <div class="card {classe}">
                    <div class="card-titulo">#{row["Ticket"]} — {titulo}</div>
                    <div class="card-meta">{badge_prio}{badge_dias}{badge_pend}</div>
                    {resumo_html}
                </div>
                ''', unsafe_allow_html=True)
                st.link_button('🔗 Abrir no Zendesk', url, use_container_width=False)
                st.markdown('<div style="margin-bottom:4px"></div>', unsafe_allow_html=True)

        if not df_prio3.empty:
            st.markdown(f'### 🚨 Prioridades ({len(df_prio3)})')
            render_cards(df_prio3, 'card-prio')
            st.markdown('---')
        if not df_atraso3.empty:
            st.markdown(f'### ⏰ Em atraso ({len(df_atraso3)})')
            render_cards(df_atraso3, 'card-atraso')
            st.markdown('---')
        if not df_normal3.empty:
            st.markdown(f'### 📂 Demais abertos ({len(df_normal3)})')
            render_cards(df_normal3, 'card-normal')

# ════════════ ABA 4 — PRIORIDADES ════════════
with aba4:
    df_p = df[df['PrioAtiva']].copy()
    p1,p2,p3,p4 = st.columns(4)
    p1.metric('🚨 Total prioridades', len(df_p)); p2.metric('📂 Em aberto', int((~df_p['Resolvido']).sum()))
    p3.metric('✅ Resolvidos', int(df_p['Resolvido'].sum())); p4.metric('⏰ Atraso +7d', int(df_p['Atrasado'].sum()))
    st.markdown('---')
    with st.expander('📊 Gráficos', expanded=True):
        col_g, col_h = st.columns(2)
        with col_g:
            st.markdown('### Por Representante')
            df_pr = df_p.groupby('RepNome').size().reset_index(name='Prioridades').sort_values('Prioridades', ascending=True)
            if not df_pr.empty: st.plotly_chart(bar_chart_pct(df_pr,'Prioridades','RepNome','#8b5cf6',height=280), use_container_width=True)
        with col_h:
            st.markdown('### Por Gerente')
            df_pg = df_p.groupby('GerNome').size().reset_index(name='Prioridades').sort_values('Prioridades', ascending=True)
            if not df_pg.empty: st.plotly_chart(bar_chart_pct(df_pg,'Prioridades','GerNome','#ef4444',height=280), use_container_width=True)
    with st.expander(f'📋 Tabela detalhada ({len(df_p)})', expanded=True):
        if df_p.empty: st.success('Nenhuma prioridade ativa.')
        else: mostrar_tabela(preparar_tabela(df_p), 'prio_det', 'Prioridades Ativas')
    st.markdown('---')
    st.markdown('### 📱 Visão Resumida — Print / WhatsApp')
    hoje_fmt = datetime.now().strftime('%d/%m/%Y')
    if df_p.empty:
        st.success('Nenhuma prioridade ativa no momento.')
    else:
        grupos = df_p.groupby('RepNome')
        linhas_html = ''
        for rep_nome, df_grp in grupos:
            linhas_html += f'<tr style="background:#C0392B"><td colspan="5" style="padding:5px 8px;color:#fff;font-weight:700;font-size:11px">👤 {rep_nome} — {len(df_grp)} ticket(s)</td></tr>'
            for _, row in df_grp.sort_values('DiasAberto', ascending=False).iterrows():
                cor_dias = '#dc2626' if row['DiasAberto'] >= DIAS_ATRASO else '#d97706'
                linhas_html += f'<tr style="background:#fff8f8;border-bottom:1px solid #f5e0e0"><td style="padding:4px 8px;font-size:11px;font-weight:600;white-space:nowrap">#{row["Ticket"]}</td><td style="padding:4px 8px;font-size:11px;max-width:220px">🚨 {str(row["Titulo"])[:55]}{"…" if len(str(row["Titulo"]))>55 else ""}</td><td style="padding:4px 8px;font-size:11px;color:{cor_dias};font-weight:700;white-space:nowrap">⏰ {row["DiasAberto"]}d</td><td style="padding:4px 8px;font-size:11px;color:#555;white-space:nowrap">{row["PendenteCom"] or "—"}</td><td style="padding:4px 8px;font-size:11px;color:#777">{str(row["Observacao"])[:35] if row["Observacao"] else "—"}</td></tr>'
        st.markdown(f'<div style="background:#fff;border:1px solid #e5e7eb;border-radius:8px;padding:0;overflow:hidden;max-width:700px"><div style="background:#C0392B;color:#fff;padding:8px 12px;font-size:13px;font-weight:700">🚨 PRIORIDADES EM ABERTO — {hoje_fmt} — {len(df_p)} ticket(s)</div><table style="width:100%;border-collapse:collapse"><thead><tr style="background:#f9f9f9;border-bottom:2px solid #e5e7eb"><th style="padding:5px 8px;font-size:10px;color:#6b7a99;text-align:left">TICKET</th><th style="padding:5px 8px;font-size:10px;color:#6b7a99;text-align:left">TÍTULO</th><th style="padding:5px 8px;font-size:10px;color:#6b7a99;text-align:left">DIAS</th><th style="padding:5px 8px;font-size:10px;color:#6b7a99;text-align:left">PENDENTE</th><th style="padding:5px 8px;font-size:10px;color:#6b7a99;text-align:left">OBSERVAÇÃO</th></tr></thead><tbody>{linhas_html}</tbody></table></div>', unsafe_allow_html=True)
        st.markdown(' ')
        with st.expander('📋 Copiar texto para WhatsApp'):
            texto_wpp = f"🚨 *PRIORIDADES EM ABERTO — {hoje_fmt}*\n\n"
            for rep_nome, df_grp in grupos:
                texto_wpp += f"👤 *{rep_nome}* ({len(df_grp)} ticket{'s' if len(df_grp)>1 else ''})\n"
                for _, row in df_grp.sort_values('DiasAberto', ascending=False).iterrows():
                    texto_wpp += f"  • #{row['Ticket']} — {str(row['Titulo'])[:50]}\n    ⏰ {row['DiasAberto']} dias | {row['PendenteCom'] or '—'}\n"
                texto_wpp += "\n"
            st.code(texto_wpp, language=None)

# ════════════ ABA 5 — REPORT SEMANAL (admin) ════════════
if aba5 is not None:
    with aba5:
        st.markdown('### 📧 Report Semanal por Representante')
        reps_report = sorted(df_raw['RepNome'].dropna().unique().tolist())
        rep_report  = st.selectbox('Selecione o representante', reps_report, key='rep_report')
        df_rep_base = df_raw[df_raw['RepNome'] == rep_report].copy()
        df_rep_ab   = df_rep_base[~df_rep_base['Resolvido']]
        ab_r  = len(df_rep_ab); at_r = int(df_rep_ab['Atrasado'].sum())
        pri_r = int(df_rep_ab['PrioAtiva'].sum()); pc_r = int((df_rep_ab['PendenteCom']=='Cliente').sum())
        ger_r = df_rep_base['GerNome'].iloc[0] if not df_rep_base.empty else '—'
        hoje  = datetime.now().strftime('%d/%m/%Y')
        st.markdown(f"""
        <div style="background:#fff;border:1px solid #e5e7eb;border-radius:8px;padding:16px;margin-bottom:16px">
            <div style="font-size:15px;font-weight:700;color:#1a2540;margin-bottom:3px">REPORT SEMANAL — PÓS-VENDAS PENDENTES REP. {rep_report.upper()}</div>
            <div style="font-size:11px;color:#6b7a99;margin-bottom:12px">Gerente: {ger_r or '—'} &nbsp;·&nbsp; Atualização: {hoje}</div>
            <div style="display:flex;gap:10px;flex-wrap:wrap">
                <div style="border:1px solid #fed7aa;border-radius:6px;padding:8px 16px;text-align:center"><div style="font-size:9px;color:#ea580c;text-transform:uppercase;font-weight:600">Abertos</div><div style="font-size:26px;font-weight:700;color:#ea580c">{ab_r}</div></div>
                <div style="border:1px solid #fecaca;border-radius:6px;padding:8px 16px;text-align:center"><div style="font-size:9px;color:#dc2626;text-transform:uppercase;font-weight:600">Atraso</div><div style="font-size:26px;font-weight:700;color:#dc2626">{at_r}</div></div>
                <div style="border:1px solid #e9d5ff;border-radius:6px;padding:8px 16px;text-align:center"><div style="font-size:9px;color:#7e22ce;text-transform:uppercase;font-weight:600">Prioridades</div><div style="font-size:26px;font-weight:700;color:#7e22ce">{pri_r}</div></div>
                <div style="border:1px solid #a5f3fc;border-radius:6px;padding:8px 16px;text-align:center"><div style="font-size:9px;color:#0891b2;text-transform:uppercase;font-weight:600">Pend. Cliente</div><div style="font-size:26px;font-weight:700;color:#0891b2">{pc_r}</div></div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown('#### 📝 Texto do resumo (copie para o e-mail)')
        texto_resumo = (
            f"Boa tarde!\n\n@{rep_report},\nSegue o status report dos pós-vendas em aberto pendentes de resolução.\n"
            f"Em anexo, segue o detalhamento dos tickets + observações e, abaixo, o resumo para acompanhamento:\n\n"
            f"Resumo dos pós-vendas em aberto:\nData da atualização: {hoje}\n\n"
            f"• {ab_r} tickets com pós-vendas em aberto;\n• {at_r} tickets atrasados, considerando tickets com mais de 7 dias em aberto;\n"
            f"• {pri_r} tickets prioritários, referente às solicitações de prorrogação de boletos;\n"
            f"• {pc_r} tickets pendentes com o cliente, por gentileza, verificar."
        )
        st.code(texto_resumo, language=None)
        st.markdown('---')
        df_cli_r = df_rep_ab[df_rep_ab['PendenteCom']=='Cliente'].sort_values('DiasAberto', ascending=False)
        st.markdown(f'#### 👤 Pós-vendas pendentes com o Cliente ({len(df_cli_r)})')
        if df_cli_r.empty: st.success('Nenhum ticket pendente com o cliente.')
        else: mostrar_tabela(preparar_tabela(df_cli_r), 'report_cli', f'Pendentes Cliente — {rep_report}')
        st.markdown('---')
        df_pos_r = df_rep_ab[df_rep_ab['PendenteCom'].str.strip().str.lower().str.contains('p.s-venda', na=False, regex=True)].sort_values('DiasAberto', ascending=False)
        st.markdown(f'#### 🏢 Pós-vendas pendentes com Pós-venda/Comercial ({len(df_pos_r)})')
        if df_pos_r.empty: st.success('Nenhum ticket pendente com Pós-venda/Comercial.')
        else: mostrar_tabela(preparar_tabela(df_pos_r), 'report_pos', f'Pendentes Pós-Venda — {rep_report}')
        st.markdown('---')
        mask_cli_r = df_rep_ab['PendenteCom'].str.strip().str.lower() == 'cliente'
        mask_pos_r = df_rep_ab['PendenteCom'].str.strip().str.lower().str.contains('p.s-venda', na=False, regex=True)
        df_outros_r = df_rep_ab[~mask_cli_r & ~mask_pos_r].sort_values('DiasAberto', ascending=False)
        if not df_outros_r.empty:
            st.markdown(f'#### 📋 Outros pendentes ({len(df_outros_r)})')
            mostrar_tabela(preparar_tabela(df_outros_r), 'report_outros', f'Outros — {rep_report}')
            st.markdown('---')
        st.markdown('#### ⬇️ Exportar Excel para anexar no e-mail')
        excel_data = gerar_excel_report(df_cli_r, df_pos_r, rep_report)
        st.download_button(
            label=f'📥 Baixar Report — {rep_report} — {hoje.replace("/","-")}.xlsx',
            data=excel_data, file_name=f'Report_{rep_report.replace(" ","_")}_{hoje.replace("/","-")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True
        )

# ════════════ ABA 6 — PLANILHA COMPLETA (admin) ════════════
if aba6 is not None:
    with aba6:
        st.markdown('### 🗂️ Planilha Completa')
        st.caption('Visualização completa de todos os dados — equivalente ao Google Sheets')
        fc1, fc2, fc3, fc4 = st.columns(4)
        with fc1:
            busca_pl = st.text_input('🔍 Buscar', placeholder='NF, cliente, ticket...', key='busca_planilha')
        with fc2:
            reps_pl = ['Todos'] + sorted(df_raw['RepNome'].dropna().unique().tolist())
            rep_pl  = st.selectbox('Representante', reps_pl, key='pl_rep')
        with fc3:
            status_pl_opts = ['Todos','Somente Abertos','Somente Resolvidos'] + sorted(df_raw['Status'].dropna().unique().tolist())
            status_pl = st.selectbox('Status', status_pl_opts, key='pl_status')
        with fc4:
            pend_pl_opts = ['Todos'] + sorted(df_raw['PendenteCom'].replace('',pd.NA).dropna().unique().tolist())
            pend_pl = st.selectbox('Pendente Com', pend_pl_opts, key='pl_pend')
        df_pl = df_raw.copy()
        if rep_pl != 'Todos': df_pl = df_pl[df_pl['RepNome'] == rep_pl]
        if status_pl == 'Somente Abertos': df_pl = df_pl[~df_pl['Resolvido']]
        elif status_pl == 'Somente Resolvidos': df_pl = df_pl[df_pl['Resolvido']]
        elif status_pl != 'Todos': df_pl = df_pl[df_pl['Status'] == status_pl]
        if pend_pl != 'Todos': df_pl = df_pl[df_pl['PendenteCom'] == pend_pl]
        if busca_pl:
            bp = busca_pl.lower()
            df_pl = df_pl[
                df_pl['Ticket'].str.lower().str.contains(bp, na=False) |
                df_pl['Titulo'].str.lower().str.contains(bp, na=False) |
                df_pl['NF'].str.lower().str.contains(bp, na=False) |
                df_pl['RazaoSocial'].str.lower().str.contains(bp, na=False) |
                df_pl['Observacao'].str.lower().str.contains(bp, na=False) |
                df_pl['ResumoIA'].str.lower().str.contains(bp, na=False) |
                df_pl['CodCliente'].str.lower().str.contains(bp, na=False)
            ]
        st.caption(f'Exibindo **{len(df_pl)}** de **{len(df_raw)}** registros')
        COLS_PL  = ['Ticket','Titulo','Representante','Gerente','DataAbertura','DataAtualizacao',
                    'HoraAtualizacao','DiasAberto','Status','PendenteCom','Motivo','Observacao',
                    'CodCliente','RazaoSocial','NF','Pedido','QuemRespondeu','UltimaAtualizacao','URL','ResumoIA']
        LABELS_PL = ['Ticket','Título','Representante','Gerente','Abertura','Últ. Data',
                     'Últ. Hora','Dias','Status','Pendente Com','Motivo','Observação',
                     'Cód. Cliente','Razão Social','NF','Pedido','Quem Respondeu','Atualização','Link','Resumo IA']
        df_show_pl = df_pl[[c for c in COLS_PL if c in df_pl.columns]].copy()
        df_show_pl.columns = LABELS_PL[:len(df_show_pl.columns)]
        df_show_pl = df_show_pl.sort_values('Dias', ascending=False)
        st.dataframe(df_show_pl, use_container_width=True, hide_index=True, height=600,
                     column_config={
                         'Dias':      st.column_config.NumberColumn('Dias', format='%d'),
                         'Link':      st.column_config.LinkColumn('Link', display_text='Abrir'),
                         'Resumo IA': st.column_config.TextColumn('Resumo IA', width='large'),
                     })
        hoje_pl = datetime.now().strftime('%d/%m/%Y')
        excel_pl = gerar_excel_tabela(df_show_pl, nome_aba='Planilha Completa', titulo=f'Planilha Completa Pós-Vendas — {hoje_pl}')
        st.download_button(
            f'📥 Exportar Excel ({len(df_pl)} registros)', data=excel_pl,
            file_name=f'Planilha_Completa_{datetime.now().strftime("%d-%m-%Y")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True
        )

st.markdown('---')
st.caption('Simplifique Representações · Dados via Google Sheets')
