from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import json, os, shutil, subprocess, re, tempfile
from datetime import datetime
from docx import Document
import openpyxl

app = Flask(__name__)
CORS(app)

TEMPLATES_DIR = '/app/templates'
SCRIPTS_DIR = '/app/scripts'

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
            return send_file(out, as_attachment=True, download_name=f'{slug}_sow.docx',
                           mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        elif ftype == 'proposal':
            out = f'{tmp}/{slug}_proposal.pptx'
            build_proposal(data, out, tmp)
            return send_file(out, as_attachment=True, download_name=f'{slug}_proposal.pptx',
                           mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        elif ftype == 'estimate':
            out = f'{tmp}/{slug}_estimate.xlsx'
            build_estimate(data, out)
            return send_file(out, as_attachment=True, download_name=f'{slug}_estimate.xlsx',
                           mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        elif ftype == 'gantt':
            out = f'{tmp}/{slug}_gantt.xlsx'
            build_gantt(data, out)
            return send_file(out, as_attachment=True, download_name=f'{slug}_gantt.xlsx',
                           mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        return jsonify({'error': 'unknown type'}), 400


def replace_in_doc(doc, old, new):
    for para in doc.paragraphs:
        for run in para.runs:
            if old.lower() in run.text.lower():
                run.text = run.text.replace(old, new)
                run.text = run.text.replace(old.lower(), new)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if old in run.text:
                            run.text = run.text.replace(old, new)


def set_cell(table, row_idx, text):
    if row_idx >= len(table.rows) or not text: return
    cell = table.rows[row_idx].cells[0]
    for p in cell.paragraphs:
        for r in p.runs: r.text = ''
    if cell.paragraphs[0].runs:
        cell.paragraphs[0].runs[0].text = text
    else:
        cell.paragraphs[0].add_run(text)


def build_sow(data, out):
    doc = Document(f'{TEMPLATES_DIR}/sow.docx')
    client = data.get('clientName', 'CLIENT')
    project = data.get('projectName', 'PROJECT')
    budget = data.get('budget', '0')
    timeline = data.get('timeline', 'TBC')
    sow = data.get('sowOutput', '')
    today = datetime.today().strftime('%-d %B %Y')

    # Replace ALL old client/project references throughout entire document
    all_replacements = [
        ('Blue Square Marketing Limited', client),
        ('Blue Square', client),
        ('AccountsPayable@bluesquare.uk.com', f'accounts@{client.lower().replace(" ","")}.com'),
        ('bluesquare.uk.com', f'{client.lower().replace(" ","")}.com'),
        ('20th March 2026', today),
        ('Charlotte Cavanagh', ''),
        ('Samsung 2026 Q2 Contact Centre Roadshow', project),
        ('Samsung', client),
        ('roadshow', 'this project'),
        ('Roadshow', 'This Project'),
    ]
    for old, new in all_replacements:
        replace_in_doc(doc, old, new)

    t = doc.tables[0]
    sec = parse_sow(sow, data)

    # Overwrite ALL content rows with correct project data
    rmap = {
        3: sec.get('summary', f'Cyphr will deliver {project} for {client}.'),
        5: sec.get('objectives', f'Deliver {project} on time and within budget of £{fmt(budget)}.'),
        7: sec.get('assumptions', f'Client to provide all required content and access within agreed timelines.\nAll third-party integrations and API access to be arranged by {client} prior to kick-off.'),
        9: sec.get('responsibilities', f'Cyphr will be responsible for all design, build and delivery activities.\n{client} will be responsible for content provision, stakeholder sign-off and UAT feedback.'),
        11: 'United Kingdom',
        15: f'Cyphr will meet with the {client} team regularly to discuss requirements and progress. Weekly status updates will be provided throughout the project.',
        17: f'Cyphr: Verity Smout\n{client}: TBC',        19: f'Cyphr will provide regular project updates during delivery. A shared project tracker will be maintained throughout.',
        23: 'Cyphr will address critical issues within 24 hours. All bugs will be tracked and resolved within agreed SLAs.',
        26: sec.get('fee', f'Fixed price of £{fmt(budget)}.'),
        28: 'Invoiced at project milestones. Payment terms: 30 days from invoice date.',
        30: f'Invoice to: Accounts Payable | {client}',
        32: 'Change requests require written approval before work commences.',
    }
    for ri, txt in rmap.items():
        set_cell(t, ri, txt)

    if len(t.rows) > 21:
        ms = sec.get('milestones', f'Week 1: Kick-off\nWeek 3: Discovery complete\nWeek {weeks(timeline)}: Build complete\nWeek {weeks(timeline)+1}: UAT\nWeek {weeks(timeline)+2}: Launch')
        set_cell(t, 21, ms)

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
            '1.3':'assumptions','assumptions':'assumptions','responsibilities':'responsibilities',
            '4.1':'milestones','milestones':'milestones','5.1':'fee','fee':'fee','3.':'fee'}
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
    project = data.get('projectName', 'PROJECT')
    budget = data.get('budget', '0')
    timeline = data.get('timeline', 'TBC')
    proposal = data.get('proposalOutput', '')
    brief = data.get('briefOutput', '')

    work = f'{tmp}/pwork/'
    result = subprocess.run(
        ['python3', f'{SCRIPTS_DIR}/unpack.py', f'{TEMPLATES_DIR}/proposal.pptx', work],
        capture_output=True, text=True, cwd=SCRIPTS_DIR
    )
    if result.returncode != 0:
        raise Exception(f'Unpack failed: {result.stderr}')

    slides = f'{work}ppt/slides/'
    paras = [p.strip() for p in (proposal or brief or '').split('\n\n') if p.strip()]
    exec_sum = paras[0][:400] if paras else f'Cyphr proposes to deliver {project} for {client}.'
    the_ask = paras[1][:300] if len(paras) > 1 else (brief[:300] if brief else f'{client} requires a strategic partner.')

    replacements = {
        'slide1.xml': [('CLIENT', client), ('PROJECT NAME', project.upper()), ('Cost Estimate ', 'Commercial Proposal ')],
        'slide2.xml': [('This proposal outlines….', exec_sum)],
        'slide3.xml': [('Activity Overview….', the_ask)],
        'slide8.xml': [
            ('Core Roadshow Experience Web App Design &amp; Build', project),
            ('Core Roadshow Experience Web App Design & Build', project),
            ('£25,314', f'£{fmt(budget)}'),
            ('£10,614', ''),
            ('2 weeks + Tour duration adhoc support', timeline),
        ],
    }
    for fname, repls in replacements.items():
        path = f'{slides}{fname}'
        if os.path.exists(path):
            c = open(path, encoding='utf-8').read()
            for old, new in repls: c = c.replace(old, new)
            open(path, 'w', encoding='utf-8').write(c)

    for i in range(1, 10):
        path = f'{slides}slide{i}.xml'
        if os.path.exists(path):
            c = open(path, encoding='utf-8').read()
            c = c.replace('CYPHR X BLUE SQUARE', f'CYPHR X {client}')
            c = c.replace('CYPHR \nX BLUE SQUARE', f'CYPHR X {client}')
            c = c.replace('PRESENTATION', 'PROPOSAL')
            open(path, 'w', encoding='utf-8').write(c)

    subprocess.run(['python3', f'{SCRIPTS_DIR}/clean.py', work], capture_output=True, cwd=SCRIPTS_DIR)
    result2 = subprocess.run(
        ['python3', f'{SCRIPTS_DIR}/pack.py', work, out, '--original', f'{TEMPLATES_DIR}/proposal.pptx'],
        capture_output=True, text=True, cwd=SCRIPTS_DIR
    )
    if not os.path.exists(out):
        raise Exception(f'Pack failed: {result2.stderr}')


def build_estimate(data, out):
    shutil.copy(f'{TEMPLATES_DIR}/estimate.xlsx', out)
    wb = openpyxl.load_workbook(out)
    ws = wb['TEMPLATE']
    client = data.get('clientName', 'CLIENT')
    project = data.get('projectName', 'PROJECT')
    today = datetime.today().strftime('%d/%m/%Y')

    # Write client/project to first available rows
    # Find row 1 and update it, or insert header info
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                if 'Cyphr Cost Estimate' in cell.value:
                    cell.value = f'Cyphr Cost Estimate — {client} / {project}'
                cell.value = (cell.value
                    .replace('CLIENT_NAME', client)
                    .replace('PROJECT_NAME', project)
                    .replace('CLIENT NAME', client)
                    .replace('PROJECT NAME', project)
                )

    # Add client info to a visible cell near the top
    ws['G1'] = f'{client} / {project}'
    ws['G2'] = f'Date: {today}'

    wb.save(out)


def build_gantt(data, out):
    shutil.copy(f'{TEMPLATES_DIR}/gantt.xlsx', out)
    wb = openpyxl.load_workbook(out)
    ws = wb.active
    client = data.get('clientName', 'CLIENT')
    project = data.get('projectName', 'PROJECT')
    label = f'{client.upper()} — {project.upper()}'

    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                cell.value = (cell.value
                    .replace('HP CRUSH 3.0 — PROJECT GANTT', label)
                    .replace('HP CRUSH 3.0', client)
                    .replace('HP Crush 3.0', client)
                    .replace('CLIENT', client)
                    .replace('PROJECT', project)
                )

    # Rename the sheet
    ws.title = f'{client} Gantt'
    wb.save(out)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
