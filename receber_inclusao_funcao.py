"""
receber_inclusao_funcao.py — versão Railway v2
===============================================
Mudanças v2:
- Protocolo sequencial DD.MM.AAAA-N (contador persistente em /data/counter.json)
- Título da tarefa: Unidade / GHE — Protocolo
- Label category22 (Inclusão de Função) + status Em andamento (percentComplete=50)
- Telefone obrigatório
- E-mail de notificação para gruposuporteengenharia@ e confirmação para o cliente
- Log de todas as submissões em /data/log.jsonl (Railway Volume)
"""

import os, time, json, fcntl, requests
from pathlib import Path
from datetime import datetime, timedelta, timezone
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

TENANT_ID     = os.environ['TENANT_ID']
CLIENT_ID     = os.environ['CLIENT_ID']
CLIENT_SECRET = os.environ['CLIENT_SECRET']
PLAN_ID       = os.environ['PLAN_ID']
BUCKET_ID     = os.environ['BUCKET_ID']

FROM_EMAIL   = 'bernardojunqueira@ocupacional.com.br'
NOTIFY_EMAIL = 'gruposuporteengenharia@ocupacional.com.br'

ASSIGNED_USERS = {
    'ac48b66a-2848-4bc9-94c6-6f2510a8c406': 'Júlia Ramos - ENG',
    '05e89169-89bf-4ac5-a532-87b9037d31ce': 'Vitor Almeida de Souza - ENG',
}

DATA_DIR    = Path('/data')
UPLOADS_DIR = Path('/tmp/uploads')
DATA_DIR.mkdir(parents=True, exist_ok=True)
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# PROTOCOLO SEQUENCIAL — contador persistente no Railway Volume
# ─────────────────────────────────────────────────────────────────────────────

COUNTER_FILE = DATA_DIR / 'counter.json'
LOCK_FILE    = DATA_DIR / 'counter.lock'

def next_protocolo() -> str:
    """Retorna próximo protocolo no formato DD.MM.AAAA-N (thread/process-safe)."""
    with open(LOCK_FILE, 'w') as lf:
        fcntl.flock(lf, fcntl.LOCK_EX)
        try:
            if COUNTER_FILE.exists():
                data = json.loads(COUNTER_FILE.read_text())
            else:
                data = {'n': 0}
            data['n'] += 1
            COUNTER_FILE.write_text(json.dumps(data))
            n = data['n']
        finally:
            fcntl.flock(lf, fcntl.LOCK_UN)

    today = datetime.now().strftime('%d.%m.%Y')
    return f'{today}-{n}'

# ─────────────────────────────────────────────────────────────────────────────
# LOG DE SUBMISSÕES — Railway Volume (/data/log.jsonl)
# ─────────────────────────────────────────────────────────────────────────────

def registrar_log(protocolo, d, task_id):
    entry = {
        'protocolo'       : protocolo,
        'data_hora'       : datetime.now().strftime('%d/%m/%Y %H:%M'),
        'cnpj'            : d.get('cnpj') or '',
        'solicitante'     : d.get('solicitante_nome', ''),
        'email'           : d.get('solicitante_email', ''),
        'telefone'        : d.get('solicitante_telefone', ''),
        'unidade'         : d.get('unidade_nome', ''),
        'setor'           : d.get('setor_nome', ''),
        'cargo'           : d.get('cargo_nome', ''),
        'ghe'             : d.get('ghe') or '',
        'task_id'         : task_id,
    }
    with open(DATA_DIR / 'log.jsonl', 'a', encoding='utf-8') as f:
        f.write(json.dumps(entry, ensure_ascii=False) + '\n')

# ─────────────────────────────────────────────────────────────────────────────
# GRAPH API — autenticação + helpers
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
# ANOTAÇÕES DA TAREFA
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
        f'Telefone       : {d.get("solicitante_telefone", "—")}',
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
    unidade = d.get('unidade_nome', '').strip()
    ghe     = d.get('ghe', '').strip()
    setor   = d.get('setor_nome', '').strip()
    loc     = ghe if ghe else setor

    titulo = f'{unidade} / {loc} — {protocolo}'
    if len(titulo) > 255: titulo = titulo[:252] + '...'

    due_dt = (datetime.now(timezone.utc) + timedelta(days=30)).strftime('%Y-%m-%dT12:00:00Z')

    assignments = {
        uid: {'@odata.type': '#microsoft.graph.plannerAssignment', 'orderHint': ' !'}
        for uid in ASSIGNED_USERS
    }

    task = gh_post('https://graph.microsoft.com/v1.0/planner/tasks', {
        'planId'            : PLAN_ID,
        'bucketId'          : BUCKET_ID,
        'title'             : titulo,
        'dueDateTime'       : due_dt,
        'percentComplete'   : 50,
        'assignments'       : assignments,
        'appliedCategories' : {'category22': True},
    })
    task_id = task['id']

    time.sleep(2)
    _, etag = gh_get(f'https://graph.microsoft.com/v1.0/planner/tasks/{task_id}/details')
    gh_patch(f'https://graph.microsoft.com/v1.0/planner/tasks/{task_id}/details',
             {'description': anotacoes}, etag)

    return task_id

# ─────────────────────────────────────────────────────────────────────────────
# E-MAILS VIA GRAPH API
# ─────────────────────────────────────────────────────────────────────────────

def enviar_email(para: list[str], assunto: str, html_body: str):
    """Envia e-mail via Graph API usando bernardojunqueira@ como remetente."""
    payload = {
        'message': {
            'subject': assunto,
            'body': {'contentType': 'HTML', 'content': html_body},
            'toRecipients': [{'emailAddress': {'address': e}} for e in para],
        },
        'saveToSentItems': False,
    }
    try:
        gh_post(f'https://graph.microsoft.com/v1.0/users/{FROM_EMAIL}/sendMail', payload)
    except Exception as e:
        print(f'[AVISO] Falha ao enviar e-mail para {para}: {e}')


def html_notificacao_interna(d, protocolo, task_id):
    cargo   = d.get('cargo_nome', '—')
    unidade = d.get('unidade_nome', '—')
    setor   = d.get('setor_nome', '—')
    ghe     = d.get('ghe') or '—'
    nome    = d.get('solicitante_nome', '—')
    email   = d.get('solicitante_email', '—')
    tel     = d.get('solicitante_telefone', '—')
    cnpj    = d.get('cnpj') or '—'
    descr   = d.get('descricao_atividade', '—').replace('\n', '<br>')

    return f"""
<div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a1a;max-width:640px;margin:0 auto">
  <div style="background:#00424b;padding:16px 24px;border-radius:6px 6px 0 0">
    <p style="color:#c4d600;font-size:11px;margin:0;text-transform:uppercase;letter-spacing:1px">Dashboard SST</p>
    <h2 style="color:#fff;margin:4px 0 0;font-size:18px">Nova Solicitação de Inclusão de Função</h2>
  </div>
  <div style="background:#e8f4f5;border:1px solid #9ec9cc;border-top:none;padding:20px 24px;border-radius:0 0 6px 6px">
    <p style="margin:0 0 16px">Uma nova solicitação foi recebida e a tarefa foi criada no Planner.</p>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <tr style="background:#00424b;color:#fff">
        <td style="padding:8px 12px;font-weight:700" colspan="2">Identificação</td>
      </tr>
      <tr style="background:#fff"><td style="padding:7px 12px;color:#555;width:38%">Protocolo</td><td style="padding:7px 12px"><strong>{protocolo}</strong></td></tr>
      <tr style="background:#f5f5f5"><td style="padding:7px 12px;color:#555">Solicitante</td><td style="padding:7px 12px">{nome}</td></tr>
      <tr style="background:#fff"><td style="padding:7px 12px;color:#555">E-mail</td><td style="padding:7px 12px">{email}</td></tr>
      <tr style="background:#f5f5f5"><td style="padding:7px 12px;color:#555">Telefone</td><td style="padding:7px 12px">{tel}</td></tr>
      <tr style="background:#fff"><td style="padding:7px 12px;color:#555">CNPJ</td><td style="padding:7px 12px">{cnpj}</td></tr>
      <tr style="background:#00424b;color:#fff">
        <td style="padding:8px 12px;font-weight:700" colspan="2">Dados do Cargo</td>
      </tr>
      <tr style="background:#fff"><td style="padding:7px 12px;color:#555">Cargo / Função</td><td style="padding:7px 12px"><strong>{cargo}</strong></td></tr>
      <tr style="background:#f5f5f5"><td style="padding:7px 12px;color:#555">Unidade</td><td style="padding:7px 12px">{unidade}</td></tr>
      <tr style="background:#fff"><td style="padding:7px 12px;color:#555">Setor</td><td style="padding:7px 12px">{setor}</td></tr>
      <tr style="background:#f5f5f5"><td style="padding:7px 12px;color:#555">GHE</td><td style="padding:7px 12px">{ghe}</td></tr>
      <tr style="background:#fff"><td style="padding:7px 12px;color:#555;vertical-align:top">Descrição</td><td style="padding:7px 12px">{descr}</td></tr>
    </table>
    <p style="margin:16px 0 0;font-size:12px;color:#555">Prazo de conclusão: <strong>30 dias corridos</strong> a partir de hoje.</p>
  </div>
  <p style="font-size:11px;color:#999;text-align:center;margin-top:12px">Grupo Ocupacional · Dashboard SST</p>
</div>"""


def html_confirmacao_cliente(d, protocolo):
    nome    = d.get('solicitante_nome', 'Cliente')
    cargo   = d.get('cargo_nome', '—')
    unidade = d.get('unidade_nome', '—')

    return f"""
<div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a1a;max-width:580px;margin:0 auto">
  <div style="background:#00424b;padding:16px 24px;border-radius:6px 6px 0 0">
    <p style="color:#c4d600;font-size:11px;margin:0;text-transform:uppercase;letter-spacing:1px">Ocupacional</p>
    <h2 style="color:#fff;margin:4px 0 0;font-size:18px">Solicitação Recebida com Sucesso</h2>
  </div>
  <div style="background:#e8f4f5;border:1px solid #9ec9cc;border-top:none;padding:24px;border-radius:0 0 6px 6px">
    <p>Olá, <strong>{nome}</strong>!</p>
    <p style="margin-top:12px">Sua solicitação de inclusão de função foi recebida pela equipe técnica da <strong>Ocupacional</strong> e já está em atendimento.</p>
    <div style="background:#fff;border:1px solid #9ec9cc;border-radius:6px;padding:16px 20px;margin:20px 0;text-align:center">
      <p style="font-size:11px;color:#4a7a7e;text-transform:uppercase;letter-spacing:1px;margin:0 0 6px">Número do Protocolo</p>
      <p style="font-size:24px;font-weight:900;color:#00424b;font-family:monospace;margin:0">{protocolo}</p>
    </div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px">
      <tr style="background:#f5f5f5"><td style="padding:7px 12px;color:#555;width:40%">Cargo / Função</td><td style="padding:7px 12px">{cargo}</td></tr>
      <tr style="background:#fff"><td style="padding:7px 12px;color:#555">Unidade</td><td style="padding:7px 12px">{unidade}</td></tr>
      <tr style="background:#f5f5f5"><td style="padding:7px 12px;color:#555">Prazo de atendimento</td><td style="padding:7px 12px"><strong>Até 30 dias corridos</strong></td></tr>
    </table>
    <p style="font-size:13px;color:#555">Em caso de dúvidas, entre em contato informando o número do protocolo acima:<br>
      <a href="mailto:suporteengenharia@ocupacional.com.br" style="color:#00424b">suporteengenharia@ocupacional.com.br</a>
    </p>
  </div>
  <p style="font-size:11px;color:#999;text-align:center;margin-top:12px">Grupo Ocupacional · Saúde e Segurança do Trabalho</p>
</div>"""

# ─────────────────────────────────────────────────────────────────────────────
# FLASK APP
# ─────────────────────────────────────────────────────────────────────────────

app = Flask(__name__)
CORS(app)

@app.route('/', methods=['GET'])
def index():
    return send_file('formulario.html')

@app.route('/health', methods=['GET'])
def health():
    counter = json.loads(COUNTER_FILE.read_text()) if COUNTER_FILE.exists() else {'n': 0}
    return jsonify({'status': 'ok', 'total_solicitacoes': counter['n']})

@app.route('/submit', methods=['POST'])
def submit():
    try:
        d = request.form.to_dict(flat=False)
        d_flat = {k: v[0] if len(v) == 1 else v for k, v in d.items()}
        if 'p7_bio_tipos' in d:
            d_flat['p7_bio_tipos'] = d['p7_bio_tipos']

        required = ['solicitante_nome', 'solicitante_email', 'solicitante_telefone',
                    'unidade_nome', 'setor_nome', 'cargo_nome', 'descricao_atividade']
        for field in required:
            if not d_flat.get(field, '').strip():
                return jsonify({'erro': f'Campo obrigatório ausente: {field}'}), 400

        protocolo = next_protocolo()

        # Salva arquivos em /tmp
        arquivos_salvos = []
        fotos = request.files.getlist('fotos')
        if fotos and fotos[0].filename:
            pasta = UPLOADS_DIR / protocolo.replace('/', '-')
            pasta.mkdir(parents=True, exist_ok=True)
            for f in fotos:
                if f.filename:
                    nome = ''.join(c for c in f.filename if c.isalnum() or c in '._- ')
                    f.save(str(pasta / nome))
                    arquivos_salvos.append(nome)

        anotacoes = formatar_anotacoes(d_flat, protocolo, arquivos_salvos)
        task_id   = criar_tarefa_planner(d_flat, protocolo, anotacoes)

        registrar_log(protocolo, d_flat, task_id)

        # E-mails (não bloqueia em caso de falha)
        enviar_email(
            [NOTIFY_EMAIL],
            f'Nova Solicitação de Inclusão de Função — {protocolo}',
            html_notificacao_interna(d_flat, protocolo, task_id),
        )
        cliente_email = d_flat.get('solicitante_email', '').strip()
        if cliente_email:
            enviar_email(
                [cliente_email],
                f'Solicitação recebida — Protocolo {protocolo}',
                html_confirmacao_cliente(d_flat, protocolo),
            )

        print(f'[OK] {protocolo} | task={task_id[:12]}...')
        return jsonify({'ok': True, 'protocolo': protocolo, 'task_id': task_id})

    except Exception as e:
        print(f'[ERRO] {e}')
        return jsonify({'erro': str(e)}), 500


if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
