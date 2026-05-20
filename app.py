from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import json, os, shutil, subprocess, tempfile, re
from datetime import datetime
from docx import Document
import openpyxl

app = Flask(__name__)
CORS(app)

TEMPLATES_DIR = '/app/templates'

def fmt(budget):
    try: return f"{int(float(budget)):,}"
    except: return str(budget)

def weeks(timeline):
    m = re.search(r'(\d+)', str(timeline))
    return int(m.group(1)) if m else 10

@app.route('/health')
def health():
    return jsonify({'status': 'ok'})

@app.route('/generate', methods=['POST', 'OPTIONS'])
def generate():
    if request.method == 'OPTIONS':
        return '', 204
    body = request.get_json()
    ftype = body.get('type')
    data = body.get('projectData', {})
    slug = re.sub(r'[^a-z0-9_]', '', (data.get('clientName') or 'project').lower().replace(' ', '_'))
    
    with tempfile.TemporaryDirectory() as tmp:
        if ftype == 'sow':
            out = f'{tmp}/{slug}_sow.docx'
            build_sow(data, out)
            return send_file(out, as_attachment=True, download_name=f'{slug}_sow.docx', mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        elif ftype == 'proposal':
            out = f'{tmp}/{slug}_proposal.pptx'
            build_proposal(data, out, tmp)
            return send_file(out, as_attachment=True, download_name=f'{slug}_proposal.pptx', mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        elif ftype == 'estimate':
            out = f'{tmp}/{slug}_estimate.xlsx'
            build_estimate(data, out)
            return send_file(out, as_attachment=True, download_name=f'{slug}_estimate.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        elif ftype == 'gantt':
            out = f'{tmp}/{slug}_gantt.xlsx'
            build_gantt(data, out)
            return send_file(out, as_attachment=True, download_name=f'{slug}_gantt.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        return jsonify({'error': 'unknown type'}), 400

def build_sow(data, out):
    doc = Document(f'{TEMPLATES_DIR}/sow.docx')
    client = data.get('clientName', 'CLIENT')
    budget = data.get('budget', '0')
    timeline = data.get('timeline', 'TBC')
    sow = data.get('sowOutput', '')
    today = datetime.today().strftime('%-d %B %Y')
    for para in doc.paragraphs:
        if 'effective from' in para.text:
            for r in para.runs: r.text = r.text.replace('20th March 2026', today)
        if 'agreement is entered' in para.text:
            for r in para.runs: r.text = r.text.replace('Blue Square Marketing Limited', client)
    t = doc.tables[0]
    sec = parse_sow(sow, data)
    rmap = {3: sec.get('summary',''), 5: sec.get('objectives',''), 7: sec.get('assumptions',''),
            9: sec.get('responsibilities',''), 11: 'United Kingdom', 17: f'Cyphr: Verity Smout\n{client}: TBC',
            26: sec.get('fee', f'Fixed price of £{fmt(budget)}.')}
    for ri, txt in rmap.items():
        if ri < len(t.rows) and txt:
            cell = t.rows[ri].cells[0]
            for p in cell.paragraphs:
                for r in p.runs: r.text = ''
            if cell.paragraphs[0].runs: cell.paragraphs[0].runs[0].text = txt
            else: cell.paragraphs[0].add_run(txt)
    if len(t.rows) > 21:
        ms = sec.get('milestones', f'Week 1: Kick-off\nWeek 3: Discovery\nWeek {weeks(timeline)}: Build complete\nWeek {weeks(timeline)+1}: UAT\nWeek {weeks(timeline)+2}: Launch')
        cell = t.rows[21].cells[0]
        for p in cell.paragraphs:
            for r in p.runs: r.text = ''
        if cell.paragraphs[0].runs: cell.paragraphs[0].runs[0].text = ms
        else: cell.paragraphs[0].add_run(ms)
    doc.save(out)

def parse_sow(text, data):
    s = {}
    if not text:
        s['summary'] = f"Cyphr will deliver {data.get('projectName','')} for {data.get('clientName','')}."
        s['fee'] = f"Fixed price of £{fmt(data.get('budget','0'))}."
        return s
    lines = text.split('\n')
    cur, buf = None, []
    keys = {'1.1':'summary','project summary':'summary','1.2':'objectives','objectives':'objectives',
            '1.3':'assumptions','assumptions':'assumptions','4.1':'milestones','milestones':'milestones',
            '5.1':'fee','fee':'fee','3.':'fee'}
    for line in lines:
        lo = line.lower().strip()
        matched = False
        for k, sec in keys.items():
            if lo.startswith(k):
                if cur and buf: s[cur] = '\n'.join(buf).strip()
                cur, buf, matched = sec, [], True
                break
        if not matched and cur and line.strip(): buf.append(line.strip())
    if cur and buf: s[cur] = '\n'.join(buf).strip()
    if 'summary' not in s: s['summary'] = text[:300]
    if 'fee' not in s: s['fee'] = f"Fixed price of £{fmt(data.get('budget','0'))}."
    return s

def build_proposal(data, out, tmp):
    client = data.get('clientName', 'CLIENT').upper()
    project = data.get('projectName', 'PROJECT').upper()
    budget = data.get('budget', '0')
    timeline = data.get('timeline', 'TBC')
    proposal = data.get('proposalOutput', '')
    brief = data.get('briefOutput', '')
    work = f'{tmp}/pwork/'
    subprocess.run(['python3', '/app/scripts/unpack.py', f'{TEMPLATES_DIR}/proposal.pptx', work], capture_output=True)
    slides = f'{work}ppt/slides/'
    paras = [p.strip() for p in proposal.split('\n\n') if p.strip()]
    exec_sum = paras[0][:400] if paras else f'Cyphr proposes to deliver {project} for {client}.'
    the_ask = paras[1][:300] if len(paras)>1 else (brief[:300] if brief else f'{client} requires a strategic partner.')
    replacements = {
        'slide1.xml': [('CLIENT', client), ('PROJECT NAME', project), ('Cost Estimate ', 'Commercial Proposal ')],
        'slide2.xml': [('This proposal outlines….', exec_sum)],
        'slide3.xml': [('Activity Overview….', the_ask)],
        'slide8.xml': [('Core Roadshow Experience Web App Design &amp; Build', data.get('projectName','PROJECT')),
                       ('£25,314', f'£{fmt(budget)}'), ('£10,614', ''), ('2 weeks + Tour duration adhoc support', timeline)],
    }
    for fname, repls in replacements.items():
        path = f'{slides}{fname}'
        if os.path.exists(path):
            c = open(path).read()
            for old, new in repls: c = c.replace(old, new)
            open(path, 'w').write(c)
    for i in range(1, 10):
        path = f'{slides}slide{i}.xml'
        if os.path.exists(path):
            c = open(path).read()
            c = c.replace('CYPHR X BLUE SQUARE', f'CYPHR X {client}')
            c = c.replace('CYPHR \nX BLUE SQUARE', f'CYPHR X {client}')
            c = c.replace('PRESENTATION', 'PROPOSAL')
            open(path, 'w').write(c)
    subprocess.run(['python3', '/app/scripts/clean.py', work], capture_output=True)
    subprocess.run(['python3', '/app/scripts/pack.py', work, out, '--original', f'{TEMPLATES_DIR}/proposal.pptx'], capture_output=True)

def build_estimate(data, out):
    shutil.copy(f'{TEMPLATES_DIR}/estimate.xlsx', out)
    wb = openpyxl.load_workbook(out)
    ws = wb['TEMPLATE']
    client = data.get('clientName', 'CLIENT')
    project = data.get('projectName', 'PROJECT')
    today = datetime.today().strftime('%d/%m/%Y')
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                cell.value = cell.value.replace('CLIENT_NAME', client).replace('PROJECT_NAME', project).replace('CLIENT NAME', client).replace('PROJECT NAME', project).replace('DATE', today)
    wb.save(out)

def build_gantt(data, out):
    shutil.copy(f'{TEMPLATES_DIR}/gantt.xlsx', out)
    wb = openpyxl.load_workbook(out)
    ws = wb.active
    client = data.get('clientName', 'CLIENT')
    project = data.get('projectName', 'PROJECT')
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                cell.value = cell.value.replace('CLIENT', client).replace('PROJECT', project)
    wb.save(out)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
