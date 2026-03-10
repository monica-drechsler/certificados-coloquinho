"""
Gerador de Certificados — Backend Flask
Rota principal: gera PPTX → PDF via LibreOffice → envia por e-mail
"""

import os, re, io, json, smtplib, zipfile, subprocess, tempfile
from pathlib import Path
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import pandas as pd
from flask import Flask, request, jsonify, send_file, render_template
from pptx import Presentation
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB

UPLOAD_FOLDER = Path(tempfile.gettempdir()) / 'certificados_uploads'
OUTPUT_FOLDER = Path(tempfile.gettempdir()) / 'certificados_output'
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)


# ══════════════════════════════════════════════════════
# ROTAS HTML
# ══════════════════════════════════════════════════════

@app.route('/')
def index():
    return render_template('index.html')


# ══════════════════════════════════════════════════════
# API: ler colunas da planilha
# ══════════════════════════════════════════════════════

@app.route('/api/planilha/colunas', methods=['POST'])
def planilha_colunas():
    if 'planilha' not in request.files:
        return jsonify({'erro': 'Nenhum arquivo enviado'}), 400

    f = request.files['planilha']
    nome = secure_filename(f.name if hasattr(f, 'name') else 'planilha')
    ext  = f.filename.rsplit('.', 1)[-1].lower()

    try:
        if ext == 'csv':
            df = pd.read_csv(f, dtype=str).fillna('')
        else:
            df = pd.read_excel(f, dtype=str).fillna('')

        df.columns = df.columns.str.strip()
        preview = df.head(5).to_dict(orient='records')

        return jsonify({
            'colunas':  list(df.columns),
            'total':    len(df),
            'preview':  preview,
        })
    except Exception as e:
        return jsonify({'erro': str(e)}), 500


# ══════════════════════════════════════════════════════
# API: gerar certificados (PPTX → PDF) + opcional e-mail
# ══════════════════════════════════════════════════════

@app.route('/api/gerar', methods=['POST'])
def gerar():
    # ── Valida arquivos ──────────────────────────────
    if 'template' not in request.files:
        return jsonify({'erro': 'Template não enviado'}), 400
    if 'planilha' not in request.files:
        return jsonify({'erro': 'Planilha não enviada'}), 400

    template_file = request.files['template']
    planilha_file = request.files['planilha']

    # Configuração (JSON no campo 'config')
    try:
        cfg = json.loads(request.form.get('config', '{}'))
    except Exception:
        return jsonify({'erro': 'Config inválida'}), 400

    seminario   = cfg.get('seminario', '')
    ministrante = cfg.get('ministrante', '')
    dia         = cfg.get('dia', '')
    mes         = cfg.get('mes', '')
    ano         = cfg.get('ano', '2026')
    carga       = cfg.get('carga', 'uma hora')
    col_nome    = cfg.get('col_nome', 'Nome completo')
    col_email   = cfg.get('col_email', '')
    enviar_email_flag = cfg.get('enviar_email', False)

    smtp_host   = cfg.get('smtp_host', 'smtp.gmail.com')
    smtp_port   = int(cfg.get('smtp_port', 587))
    smtp_user   = cfg.get('smtp_user', '')
    smtp_pass   = cfg.get('smtp_pass', '')
    smtp_nome   = cfg.get('smtp_nome', 'Coloquinho da Pós')
    assunto_tpl = cfg.get('assunto', 'Certificado — {seminario}')
    corpo_tpl   = cfg.get('corpo', CORPO_PADRAO)

    # ── Lê planilha ──────────────────────────────────
    try:
        ext = planilha_file.filename.rsplit('.', 1)[-1].lower()
        if ext == 'csv':
            df = pd.read_csv(planilha_file, dtype=str).fillna('')
        else:
            df = pd.read_excel(planilha_file, dtype=str).fillna('')
        df.columns = df.columns.str.strip()
    except Exception as e:
        return jsonify({'erro': f'Erro ao ler planilha: {e}'}), 400

    if col_nome not in df.columns:
        return jsonify({'erro': f'Coluna "{col_nome}" não encontrada. Disponíveis: {list(df.columns)}'}), 400

    # ── Salva template temporariamente ───────────────
    template_path = UPLOAD_FOLDER / secure_filename(template_file.filename or 'template.pptx')
    template_file.save(str(template_path))

    # ── Limpa outputs antigos ─────────────────────────
    for f_ in OUTPUT_FOLDER.glob('*'):
        try: f_.unlink()
        except: pass

    # ── Processa cada participante ────────────────────
    resultados = []
    pdfs_gerados = []

    for idx, row in df.iterrows():
        nome = str(row.get(col_nome, '')).strip()
        if not nome:
            resultados.append({'idx': idx+1, 'nome': '(vazio)', 'status': 'pulado', 'msg': 'Nome vazio'})
            continue

        email_dest = str(row.get(col_email, '')).strip() if col_email else ''

        try:
            # Substituições ordenadas para o template do Coloquinho
            substituicoes = [
                ('<<Nome completo>>', nome),
                (f'XX de XXXXXX de {ano}', f'{dia} de {mes} de {ano}'),
                ('"XXXXXX"',              f'"{seminario}"'),
                ('por XXXXXX',            f'por {ministrante}'),
                ('de XXXXXX de',          f'de {mes} de'),
            ]

            # Gera PPTX personalizado
            nome_arq   = nome_seguro(nome)
            pptx_path  = OUTPUT_FOLDER / f'{nome_arq}.pptx'
            pdf_path   = OUTPUT_FOLDER / f'{nome_arq}.pdf'

            gerar_pptx(str(template_path), substituicoes, str(pptx_path))

            # Converte para PDF via LibreOffice
            converter_pdf(str(pptx_path), str(OUTPUT_FOLDER))

            if not pdf_path.exists():
                raise FileNotFoundError(f'PDF não gerado: {pdf_path}')

            pdfs_gerados.append({'nome': nome, 'pdf': str(pdf_path), 'email': email_dest})

            # Envia e-mail se configurado
            msg_email = ''
            if enviar_email_flag and email_dest and smtp_user and smtp_pass:
                dados_subst = {'{nome}': nome, '{seminario}': seminario,
                               '{data}': f'{dia} de {mes} de {ano}', '{ministrante}': ministrante}
                assunto = substituir_dict(assunto_tpl, dados_subst)
                corpo   = substituir_dict(corpo_tpl,   dados_subst)
                enviar_email(smtp_host, smtp_port, smtp_user, smtp_pass, smtp_nome,
                             email_dest, assunto, corpo, str(pdf_path))
                msg_email = f' → {email_dest}'

            resultados.append({'idx': idx+1, 'nome': nome, 'status': 'ok',
                                'arquivo': nome_arq + '.pdf', 'msg': 'Gerado' + msg_email})

        except Exception as e:
            resultados.append({'idx': idx+1, 'nome': nome, 'status': 'erro', 'msg': str(e)})

    return jsonify({
        'resultados': resultados,
        'total':      len(df),
        'sucesso':    sum(1 for r in resultados if r['status'] == 'ok'),
        'erros':      sum(1 for r in resultados if r['status'] == 'erro'),
    })


# ══════════════════════════════════════════════════════
# API: baixar ZIP com todos os PDFs
# ══════════════════════════════════════════════════════

@app.route('/api/download-zip')
def download_zip():
    pdfs = list(OUTPUT_FOLDER.glob('*.pdf'))
    if not pdfs:
        return jsonify({'erro': 'Nenhum PDF encontrado. Gere os certificados primeiro.'}), 404

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for pdf in pdfs:
            zf.write(pdf, pdf.name)
    buf.seek(0)

    return send_file(buf, mimetype='application/zip',
                     as_attachment=True, download_name='certificados.zip')


# ══════════════════════════════════════════════════════
# API: preview de um certificado (retorna PDF do idx)
# ══════════════════════════════════════════════════════

@app.route('/api/preview', methods=['POST'])
def preview():
    """Gera e retorna o PDF de UM participante para preview."""
    if 'template' not in request.files:
        return jsonify({'erro': 'Template não enviado'}), 400

    template_file = request.files['template']
    try:
        cfg = json.loads(request.form.get('config', '{}'))
    except Exception:
        return jsonify({'erro': 'Config inválida'}), 400

    nome        = cfg.get('nome', 'Participante Exemplo')
    seminario   = cfg.get('seminario', 'Título do Seminário')
    ministrante = cfg.get('ministrante', 'Nome do Ministrante')
    dia         = cfg.get('dia', 'XX')
    mes         = cfg.get('mes', 'mês')
    ano         = cfg.get('ano', '2026')
    carga       = cfg.get('carga', 'uma hora')

    template_path = UPLOAD_FOLDER / 'preview_template.pptx'
    template_file.save(str(template_path))

    substituicoes = [
        ('<<Nome completo>>', nome),
        (f'XX de XXXXXX de {ano}', f'{dia} de {mes} de {ano}'),
        ('"XXXXXX"',              f'"{seminario}"'),
        ('por XXXXXX',            f'por {ministrante}'),
        ('de XXXXXX de',          f'de {mes} de'),
    ]

    try:
        pptx_path = UPLOAD_FOLDER / 'preview.pptx'
        pdf_path  = UPLOAD_FOLDER / 'preview.pdf'
        if pdf_path.exists(): pdf_path.unlink()

        gerar_pptx(str(template_path), substituicoes, str(pptx_path))
        converter_pdf(str(pptx_path), str(UPLOAD_FOLDER))

        if not pdf_path.exists():
            return jsonify({'erro': 'Falha ao gerar preview PDF'}), 500

        return send_file(str(pdf_path), mimetype='application/pdf')
    except Exception as e:
        return jsonify({'erro': str(e)}), 500


# ══════════════════════════════════════════════════════
# FUNÇÕES AUXILIARES
# ══════════════════════════════════════════════════════

def gerar_pptx(template_path: str, substituicoes: list, saida: str):
    prs = Presentation(template_path)
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                texto = ''.join(r.text for r in para.runs)
                texto_novo = texto
                for antigo, novo in substituicoes:
                    texto_novo = texto_novo.replace(antigo, novo)
                if texto != texto_novo and para.runs:
                    para.runs[0].text = texto_novo
                    for r in para.runs[1:]:
                        r.text = ''
    prs.save(saida)


def converter_pdf(pptx_path: str, pasta_saida: str):
    res = subprocess.run(
        ['libreoffice', '--headless', '--convert-to', 'pdf',
         '--outdir', pasta_saida, pptx_path],
        capture_output=True, text=True, timeout=60
    )
    if res.returncode != 0:
        raise RuntimeError(f'LibreOffice: {res.stderr[:300]}')


def enviar_email(host, port, usuario, senha, nome_rem, dest, assunto, corpo_html, pdf_path):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = assunto
    msg['From']    = f'{nome_rem} <{usuario}>'
    msg['To']      = dest
    msg.attach(MIMEText(re.sub(r'<[^>]+>', '', corpo_html), 'plain', 'utf-8'))
    msg.attach(MIMEText(corpo_html, 'html', 'utf-8'))
    with open(pdf_path, 'rb') as f:
        parte = MIMEBase('application', 'pdf')
        parte.set_payload(f.read())
    encoders.encode_base64(parte)
    parte.add_header('Content-Disposition', 'attachment', filename=Path(pdf_path).name)
    msg.attach(parte)
    with smtplib.SMTP(host, port) as smtp:
        smtp.starttls()
        smtp.login(usuario, senha)
        smtp.sendmail(usuario, dest, msg.as_string())


def substituir_dict(texto, d):
    for k, v in d.items():
        texto = texto.replace(k, v)
    return texto


def nome_seguro(nome: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', nome).strip() or 'certificado'


CORPO_PADRAO = """
<html><body style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
  <div style="background:#1a1410;color:#f7f3ec;padding:28px;border-radius:8px 8px 0 0;">
    <h2 style="margin:0;font-size:20px;">🎓 Seu certificado chegou!</h2>
  </div>
  <div style="background:#f9f6f0;padding:28px;border:1px solid #e0d8cc;border-radius:0 0 8px 8px;">
    <p>Olá, <strong>{nome}</strong>!</p>
    <p>Segue em anexo o seu certificado de participação no seminário
    <strong>"{seminario}"</strong>, realizado no <em>Coloquinho da Pós</em>
    no IME-USP em {data}.</p>
    <p>Obrigado pela sua presença! 🙌</p>
    <hr style="border:none;border-top:1px solid #e0d8cc;margin:20px 0;">
    <p style="color:#999;font-size:12px;">Organização: Raquel Mansano Gonçalves Cenciarelli</p>
  </div>
</body></html>
"""

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
