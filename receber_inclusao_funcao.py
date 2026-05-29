"""
receber_inclusao_funcao.py — versão Railway
============================================
Backend Flask para o formulário de Inclusão de Função.
Recebe POST, cria tarefa no Planner via Graph API.

Deploy: Railway (gunicorn via Procfile)
Uploads: /tmp/uploads/{protocolo}/ (ephemeral — nomes ficam nas anotações da tarefa)
"""

import os, sys, time, random, string, requests
from pathlib import Path
from datetime import datetime, timedelta, timezone
from flask import Flask, request, jsonify
from flask_cors import CORS

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO — pode sobrescrever via variáveis de ambiente no Railway
# ─────────────────────────────────────────────────────────────────────────────

TENANT_ID     = os.environ['TENANT_ID']
CLIENT_ID     = os.environ['CLIENT_ID']
CLIENT_SECRET = os.environ['CLIENT_SECRET']
PLAN_ID       = os.environ['PLAN_ID']
BUCKET_ID     = os.environ['BUCKET_ID']

ASSIGNED_USERS = {
    'ac48b66a-2848-4bc9-94c6-6f2510a8c406': 'Júlia Ramos - ENG',
    '05e89169-89bf-4ac5-a532-87b9037d31ce': 'Vitor Almeida de Souza - ENG',
}

UPLOADS_DIR = Path('/tmp/uploads')
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# GRAPH API
# ─────────────────────────────────────────────────────────────────────────────

_token_cache = {'token': '', 'expires': 0}

def get_token():
    if time.time() < _token_cache['expires'] - 60:
        return _token_cache['token']
    r = requests.post(
        f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token',
        data={'grant_type': 'client_credentials', 'client_id': CLIENT_ID,
              'client_secret': CLIENT_SECRET, 'scope': 'https://graph.microsoft.com/.default'},
        timeout=30)
    r.raise_for_status()
    d = r.json()
    _token_cache['token']   = d['access_token']
    _token_cache['expires'] = time.time() + d.get('expires_in', 3600)
    return _token_cache['token']

def gh_post(url, body):
    h = {'Authorization': f'Bearer {get_token()}', 'Content-Type': 'application/json'}
    r = requests.post(url, headers=h, json=body, timeout=30)
    r.raise_for_status()
    return r.json()

def gh_get(url):
    h = {'Authorization': f'Bearer {get_token()}'}
    r = requests.get(url, headers=h, timeout=30)
    r.raise_for_status()
    return r.json(), r.headers.get('ETag', '*')

def gh_patch(url, body, etag):
    h = {'Authorization': f'Bearer {get_token()}',
         'Content-Type': 'application/json', 'If-Match': etag}
    r = requests.patch(url, headers=h, json=body, timeout=30)
    r.raise_for_status()

# ─────────────────────────────────────────────────────────────────────────────
# FORMATAÇÃO DAS ANOTAÇÕES
# ─────────────────────────────────────────────────────────────────────────────

P5_LABELS  = {'nao': 'Não', 'sim_pequeno': 'Sim — Pequeno Porte', 'sim_grande': 'Sim — Grande Porte'}
P10_LABELS = {'sim': 'Sim', 'duvida': 'Tenho dúvida (solicitar contato do técnico)', 'nao': 'Não'}
YN = {'sim': 'Sim', 'nao': 'Não'}

def fmt_yn(v): return YN.get(v, v or '—')

def formatar_anotacoes(d, protocolo, arquivos):
    now_str = datetime.now().strftime('%d/%m/%Y %H:%M')
    linhas = [
        '=== SOLICITAÇÃO DE INCLUSÃO DE FUNÇÃO ===',
        f'Protocolo : {protocolo}',
        f'Data/hora : {now_str}',
        '',
        '--- IDENTIFICAÇÃO ---',
        f'CNPJ           : {d.get("cnpj") or "Não informado"}',
        f'Solicitante    : {d.get("solicitante_nome", "—")}',
        f'E-mail         : {d.get("solicitante_email", "—")}',
        f'Telefone       : {d.get("solicitante_telefone") or "Não informado"}',
        '',
        '--- DADOS DO CARGO ---',
        f'Unidade        : {d.get("unidade_nome", "—")}',
        f'Setor          : {d.get("setor_nome", "—")}',
        f'Cargo / Função : {d.get("cargo_nome", "—")}',
        f'GHE            : {d.get("ghe") or "Não informado"}',
        f'Descrição setor: {d.get("descricao_setor") or "Não informado"}',
        '',
        '--- DESCRIÇÃO DE ATIVIDADES ---',
        d.get('descricao_atividade', '—'),
        '',
        '--- QUESTIONÁRIO ---',
        f'P1  Head set              : {fmt_yn(d.get("p1_headset"))}',
        f'P2  Ruídos                : {fmt_yn(d.get("p2_ruidos"))}',
    ]
    if d.get('p2_ruidos') == 'sim':
        linhas.append(f'    Máquinas/equipamentos : {d.get("p2_ruidos_desc", "—")}')
    linhas += [
        f'P3  Trabalho em altura    : {fmt_yn(d.get("p3_altura"))}',
        f'P4  Espaço confinado      : {fmt_yn(d.get("p4_confinado"))}',
        f'P5  Conduz veículos       : {P5_LABELS.get(d.get("p5_veiculos",""), d.get("p5_veiculos","—"))}',
        f'P6  Produtos químicos     : {fmt_yn(d.get("p6_quimicos"))}',
    ]
    if d.get('p6_quimicos') == 'sim':
        linhas.append(f'    Descrição produtos    : {d.get("p6_quimicos_desc", "—")}')
    linhas.append(f'P7  Agentes biológicos   : {fmt_yn(d.get("p7_biologicos"))}')
    if d.get('p7_biologicos') == 'sim':
        tipos = d.get('p7_bio_tipos') or []
        if isinstance(tipos, str): tipos = [tipos]
        linhas.append(f'    Tipos exposição       : {"; ".join(tipos) if tipos else "—"}')
    linhas += [
        f'P8  Manipulação alimentos : {fmt_yn(d.get("p8_alimentos"))}',
        f'P9  EPI                   : {fmt_yn(d.get("p9_epi"))}',
    ]
    if d.get('p9_epi') == 'sim':
        linhas.append(f'    Lista EPIs/CAs        : {d.get("p9_epi_lista", "—")}')
    linhas.append(f'P10 Adicional/insalubridade: {P10_LABELS.get(d.get("p10_adicional",""), d.get("p10_adicional","—"))}')
    if d.get('p10_adicional') == 'sim':
        linhas.append(f'    Motivo                : {d.get("p10_adicional_motivo", "—")}')

    linhas += ['', '--- ARQUIVOS RECEBIDOS ---']
    if arquivos:
        linhas.append('Solicitar ao cliente caso necessário:')
        linhas += [f'  · {a}' for a in arquivos]
    else:
        linhas.append('  Nenhum arquivo enviado.')

    return '\n'.join(linhas)

# ─────────────────────────────────────────────────────────────────────────────
# CRIAR TAREFA NO PLANNER
# ─────────────────────────────────────────────────────────────────────────────

def criar_tarefa_planner(d, protocolo, anotacoes):
    cargo   = d.get('cargo_nome', 'Nova Função').strip()
    unidade = d.get('unidade_nome', '').strip()
    setor   = d.get('setor_nome', '').strip()

    titulo = f'{cargo} — {unidade} / {setor}'
    if len(titulo) > 255: titulo = titulo[:252] + '...'

    due_dt = (datetime.now(timezone.utc) + timedelta(days=30)).strftime('%Y-%m-%dT12:00:00Z')

    assignments = {
        uid: {'@odata.type': '#microsoft.graph.plannerAssignment', 'orderHint': ' !'}
        for uid in ASSIGNED_USERS
    }

    task = gh_post('https://graph.microsoft.com/v1.0/planner/tasks', {
        'planId': PLAN_ID, 'bucketId': BUCKET_ID, 'title': titulo,
        'dueDateTime': due_dt, 'assignments': assignments,
    })
    task_id = task['id']

    time.sleep(2)
    _, etag = gh_get(f'https://graph.microsoft.com/v1.0/planner/tasks/{task_id}/details')
    gh_patch(f'https://graph.microsoft.com/v1.0/planner/tasks/{task_id}/details',
             {'description': anotacoes}, etag)

    return task_id

# ─────────────────────────────────────────────────────────────────────────────
# FLASK APP
# ─────────────────────────────────────────────────────────────────────────────

app = Flask(__name__)
CORS(app)

@app.route('/', methods=['GET'])
def index():
    return jsonify({'status': 'Ocupacional — Inclusão de Função API', 'version': '1.0'})

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/submit', methods=['POST'])
def submit():
    try:
        d = request.form.to_dict(flat=False)
        d_flat = {k: v[0] if len(v) == 1 else v for k, v in d.items()}
        if 'p7_bio_tipos' in d:
            d_flat['p7_bio_tipos'] = d['p7_bio_tipos']

        required = ['solicitante_nome', 'solicitante_email',
                    'unidade_nome', 'setor_nome', 'cargo_nome', 'descricao_atividade']
        for field in required:
            if not d_flat.get(field, '').strip():
                return jsonify({'erro': f'Campo obrigatório ausente: {field}'}), 400

        sufixo    = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
        protocolo = f'INC-{datetime.now().strftime("%Y%m%d")}-{sufixo}'

        # Salva arquivos em /tmp (ephemeral — nomes ficam nas anotações)
        arquivos_salvos = []
        fotos = request.files.getlist('fotos')
        if fotos and fotos[0].filename:
            pasta = UPLOADS_DIR / protocolo
            pasta.mkdir(parents=True, exist_ok=True)
            for f in fotos:
                if f.filename:
                    nome = ''.join(c for c in f.filename if c.isalnum() or c in '._- ')
                    f.save(str(pasta / nome))
                    arquivos_salvos.append(nome)

        anotacoes = formatar_anotacoes(d_flat, protocolo, arquivos_salvos)
        task_id   = criar_tarefa_planner(d_flat, protocolo, anotacoes)

        print(f'[OK] protocolo={protocolo} task={task_id[:12]}...')
        return jsonify({'ok': True, 'protocolo': protocolo, 'task_id': task_id})

    except Exception as e:
        print(f'[ERRO] {e}')
        return jsonify({'erro': str(e)}), 500


if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
