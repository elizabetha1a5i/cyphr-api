from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import json, os, re, tempfile, requests
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import shutil, subprocess

app = Flask(__name__)
CORS(app)

TEMPLATES_DIR = '/app/templates'
SCRIPTS_DIR   = '/app/scripts'

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIG — single source of truth for team, rates, brand
# ═══════════════════════════════════════════════════════════════════════════════

CYPHR_TEAM = {
    'strategy_rob':   {'name': 'Strategy Lead (Rob)',       'rate': 950},
    'strategy_james': {'name': 'Strategy Lead (James)',     'rate': 950},
    'tech_lead':      {'name': 'Tech Lead (Redian)',        'rate': 700},
    'tech_lead_p2':   {'name': 'Tech Lead (Redian)',        'rate': 500},
    'producer':       {'name': 'Producer (Verity)',         'rate': 500},
    'qa':             {'name': 'QA',                        'rate': 450},
    'support':        {'name': 'Support',                   'rate': 500},
    'hosting':        {'name': 'Hosting & 3rd Party',       'rate': 500},
    'dev_mid':        {'name': 'Developer (Mid)',           'rate': 200},
    'dev_senior':     {'name': 'Developer (Senior)',        'rate': 300},
}

CYPHR_DELIVERY_LEAD      = 'Verity Smout'
CYPHR_DELIVERY_LEAD_ROLE = 'Head of Production'
CYPHR_EMAIL              = 'hello@cyphr.studio'
CYPHR_SITE               = 'cyphr.studio'

# Brand colours
C_PURPLE  = RGBColor(0x5B, 0x4F, 0xD9)
C_DARK    = RGBColor(0x1F, 0x1F, 0x1F)
C_GREY    = RGBColor(0x6B, 0x72, 0x80)
C_LIGHT   = RGBColor(0xE8, 0xE5, 0xFF)
C_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)

# ═══════════════════════════════════════════════════════════════════════════════
# AI CALL — model-switchable via env var
# ═══════════════════════════════════════════════════════════════════════════════

def call_ai(prompt, max_tokens=2000):
    """
    Call the configured AI model. Switch by setting AI_PROVIDER env var:
      anthropic  (default) — uses ANTHROPIC_API_KEY
      openai               — uses OPENAI_API_KEY
      gemini               — uses GEMINI_API_KEY
    """
    provider = os.environ.get('AI_PROVIDER', 'anthropic').lower()

    if provider == 'openai':
        key = os.environ.get('OPENAI_API_KEY', '')
        res = requests.post('https://api.openai.com/v1/chat/completions',
            headers={'Authorization': f'Bearer {key}', 'Content-Type': 'application/json'},
            json={'model': 'gpt-4o-mini', 'max_tokens': max_tokens,
                  'messages': [{'role': 'user', 'content': prompt}]}, timeout=60)
        res.raise_for_status()
        return res.json()['choices'][0]['message']['content']

    elif provider == 'gemini':
        key = os.environ.get('GEMINI_API_KEY', '')
        res = requests.post(
            f'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={key}',
            headers={'Content-Type': 'application/json'},
            json={'contents': [{'parts': [{'text': prompt}]}]}, timeout=60)
        res.raise_for_status()
        return res.json()['candidates'][0]['content']['parts'][0]['text']

    else:  # anthropic (default)
        key = os.environ.get('ANTHROPIC_API_KEY', '')
        res = requests.post('https://api.anthropic.com/v1/messages',
            headers={'x-api-key': key, 'anthropic-version': '2023-06-01',
                     'Content-Type': 'application/json'},
            json={'model': 'claude-haiku-4-5-20251001', 'max_tokens': max_tokens,
                  'messages': [{'role': 'user', 'content': prompt}]}, timeout=60)
        res.raise_for_status()
        return res.json()['content'][0]['text']


def parse_json_response(text):
    """Strip markdown fences and parse JSON from AI response."""
    clean = re.sub(r'^```(?:json)?\s*', '', text.strip(), flags=re.MULTILINE)
    clean = re.sub(r'\s*```$', '', clean.strip(), flags=re.MULTILINE)
    return json.loads(clean.strip())


# ═══════════════════════════════════════════════════════════════════════════════
# PLACEHOLDER CONVENTION
# All missing data surfaces as [CONFIRM: reason] — searchable before sending
# ═══════════════════════════════════════════════════════════════════════════════

def PH(reason):
    return f'[CONFIRM: {reason}]'


# ═══════════════════════════════════════════════════════════════════════════════
# VALIDATION + VERIFICATION
# ═══════════════════════════════════════════════════════════════════════════════

REQUIRED = {
    'brief':    ['clientName', 'projectName'],
    'sow':      ['clientName', 'projectName', 'sowOutput'],
    'proposal': ['clientName', 'projectName', 'proposalOutput'],
    'estimate': ['clientName', 'projectName'],
    'gantt':    ['clientName', 'projectName'],
}

def validate(ftype, data):
    missing = []
    for f in REQUIRED.get(ftype, []):
        if not str(data.get(f, '')).strip():
            missing.append(f)
    if missing:
        return False, {'error': 'Missing required fields', 'fields': missing}
    return True, []

def verify_output(doc_type, content, data):
    """
    Run the verifier against generated content before rendering.
    Returns list of issues — empty means clean.
    """
    issues = []

    # 1. Check for leftover placeholder text from previous projects
    bleed_terms = [
        'Blue Square', 'Charlotte Cavanagh', 'charlotte.cavanagh',
        'Samsung Roadshow', 'Contact Centre Roadshow', 'photo booth',
        'HP Local Heroes', 'MentorLink', 'Spotify Kids', 'Google Chromecast',
        'AccountsPayable@bluesquare', 'Tate House, Watermark Way',
        'lorem ipsum', 'Lorem ipsum', 'morem ipsum',
        'Fantasy Golf', "Rick's Roll up", 'Birdie Sauce',
    ]
    content_str = json.dumps(content) if isinstance(content, dict) else str(content)
    for term in bleed_terms:
        if term.lower() in content_str.lower():
            issues.append(f'Previous project data found: "{term}" — AI may have used template content')

    # 2. Check for invented person names (not in team config or source data)
    known = {v['name'].lower() for v in CYPHR_TEAM.values()}
    known.add(CYPHR_DELIVERY_LEAD.lower())
    source = ' '.join(filter(None, [
        data.get('clientName',''), data.get('projectName',''),
        data.get('requirements',''), data.get('bgNotes',''),
        data.get('briefOutput',''), data.get('sowOutput',''),
    ])).lower()
    skip_words = {
        'project','phase','design','build','launch','sprint','discovery',
        'planning','review','sign','off','kick','total','fixed','price',
        'united','kingdom','north','south','east','west','new','old',
    }
    name_re = re.compile(r'\b([A-Z][a-z]{2,} [A-Z][a-z]{2,})\b')
    for name in set(name_re.findall(content_str)):
        words = name.lower().split()
        if any(w in skip_words for w in words): continue
        if name.lower() in known: continue
        if name.lower() in source: continue
        issues.append(f'Possible invented name: "{name}" — not in team config or source data')

    # 3. Check budget figures don't exceed project budget
    raw = str(data.get('budget','')).replace('£','').replace(',','').strip()
    if raw and raw not in ('0',''):
        try:
            user_budget = int(float(raw))
            for fee in [int(m.replace(',','')) for m in re.findall(r'£([\d,]+)', content_str)]:
                if fee > user_budget * 1.2 and fee > 1000:
                    issues.append(f'Fee £{fee:,} exceeds project budget £{user_budget:,} — check this figure')
        except: pass

    # 4. Check required fields are actually populated in output
    if isinstance(content, dict):
        for key, val in content.items():
            if isinstance(val, str) and val.strip() in ('', 'N/A', 'TBC', 'None'):
                issues.append(f'Field "{key}" is empty in generated output — missing from context')

    return issues


# ═══════════════════════════════════════════════════════════════════════════════
# DOCUMENT SPECS — what the AI receives per document type
# ═══════════════════════════════════════════════════════════════════════════════

def sow_spec(data):
    client   = data.get('clientName', '')
    project  = data.get('projectName', '')
    budget   = data.get('budget', '')
    timeline = data.get('timeline', '')
    sow_text = data.get('sowOutput', '')
    brief    = data.get('briefOutput', '')
    estimate = data.get('estimateOutput', '')
    today    = datetime.today().strftime('%-d %B %Y')

    team_names = ', '.join(v['name'] for v in CYPHR_TEAM.values())

    prompt = f"""You are generating a Statement of Work for Cyphr Studio. Return ONLY valid JSON, no markdown, no explanation.

PROJECT DATA:
Client: {client}
Project: {project}
Budget: £{budget}
Timeline: {timeline}
Brief: {brief[:800] if brief else 'Not provided'}
AI SOW draft: {sow_text[:1200] if sow_text else 'Not provided'}
Estimate: {estimate[:600] if estimate else 'Not provided'}
Today: {today}

RULES:
- Only use names from this list for Cyphr team: {team_names}. Cyphr delivery lead is always: {CYPHR_DELIVERY_LEAD}.
- Client contact is unknown — use exactly the string: [CONFIRM: client project lead name and email]
- Budget: if unknown use exactly: [CONFIRM: budget not yet confirmed]
- Milestones: derive from timeline and project type. Format as "Milestone name — Date" per line. If dates unknown use: [CONFIRM: milestone dates — confirm at kick-off]
- Fee: state as fixed price total matching the budget. If budget is 0 use: [CONFIRM: fee — confirm with estimate]
- Do NOT invent addresses, company numbers, or contact details
- Keep each section focused and professional. No filler.

Return this exact JSON structure:
{{
  "effective_date": "{today}",
  "client": "{client}",
  "project": "{project}",
  "summary": "2-3 sentences describing what Cyphr will deliver and for whom",
  "objectives": "bullet points — key project objectives and outputs",
  "assumptions": "bullet points — what client must provide, key dependencies",
  "responsibilities": "Cyphr responsibilities on one side, client responsibilities on the other",
  "location": "United Kingdom",
  "milestones": "one milestone per line as: Name — Date",
  "fee": "fixed price fee statement with total amount",
  "payment_terms": "invoice and payment schedule",
  "client_contact": "[CONFIRM: client project lead name and email]",
  "delivery_lead": "{CYPHR_DELIVERY_LEAD}",
  "delivery_lead_role": "{CYPHR_DELIVERY_LEAD_ROLE}"
}}"""
    return prompt

def proposal_spec(data):
    client   = data.get('clientName', '')
    project  = data.get('projectName', '')
    budget   = data.get('budget', '')
    timeline = data.get('timeline', '')
    sector   = data.get('sector', '')
    proposal = data.get('proposalOutput', '')
    brief    = data.get('briefOutput', '')
    today    = datetime.today().strftime('%-d %B %Y')

    prompt = f"""You are generating a commercial proposal for Cyphr Studio. Return ONLY valid JSON, no markdown, no explanation.

PROJECT DATA:
Client: {client}
Project: {project}
Sector: {sector}
Budget: £{budget}
Timeline: {timeline}
Brief: {brief[:800] if brief else 'Not provided'}
Proposal draft: {proposal[:1000] if proposal else 'Not provided'}

RULES:
- Write as Cyphr Studio — confident, direct, no fluff
- executive_summary: 2-3 sentences max
- opportunity: what problem this solves for the client
- approach: how Cyphr will tackle it — 3-4 bullet points
- deliverables: what the client receives — specific, not vague
- why_cyphr: 2-3 sentences on why Cyphr is the right partner for this
- investment: fee summary with total matching budget
- Do NOT reference Samsung, Blue Square, roadshow, or any other client

Return this exact JSON:
{{
  "client": "{client}",
  "project": "{project}",
  "date": "{today}",
  "executive_summary": "...",
  "opportunity": "...",
  "approach": ["bullet 1", "bullet 2", "bullet 3"],
  "deliverables": ["deliverable 1", "deliverable 2", "deliverable 3"],
  "why_cyphr": "...",
  "investment": "...",
  "timeline": "{timeline}"
}}"""
    return prompt

def estimate_spec(data):
    client   = data.get('clientName', '')
    project  = data.get('projectName', '')
    budget   = data.get('budget', '0')
    timeline = data.get('timeline', '')
    estimate = data.get('estimateOutput', '')
    brief    = data.get('briefOutput', '')

    try:
        total = int(float(str(budget).replace('£','').replace(',','')))
    except:
        total = 0

    team_list = '\n'.join(f"  {k}: {v['name']} @ £{v['rate']}/day" for k,v in CYPHR_TEAM.items())

    prompt = f"""You are building a cost estimate for a Cyphr Studio project. Return ONLY valid JSON, no markdown.

PROJECT DATA:
Client: {client}
Project: {project}
Total budget: £{total:,}
Timeline: {timeline}
Brief: {brief[:600] if brief else 'Not provided'}
Estimate notes: {estimate[:600] if estimate else 'Not provided'}

AVAILABLE TEAM (use ONLY these, pick what fits the project):
{team_list}

RULES:
- Phases must add up to approximately £{total:,} total
- Only include team members that make sense for this project type
- Days must be realistic for the timeline and project type
- Each phase has a name, list of roles with days, and subtotal
- Use the exact rate values from the team list above — do not invent rates
- If budget is 0, estimate based on project type and timeline

Return this exact JSON:
{{
  "client": "{client}",
  "project": "{project}",
  "total_budget": {total},
  "phases": [
    {{
      "name": "Phase name",
      "roles": [
        {{"team_key": "strategy_rob", "days": 2}},
        {{"team_key": "producer", "days": 3}}
      ]
    }}
  ],
  "assumptions": ["assumption 1", "assumption 2"],
  "exclusions": ["exclusion 1", "exclusion 2"]
}}"""
    return prompt


# ═══════════════════════════════════════════════════════════════════════════════
# DOCUMENT BUILDERS — from JSON spec, not from templates
# ═══════════════════════════════════════════════════════════════════════════════

def build_sow(data, out):
    # 1. Generate structured content via AI
    spec_json = call_ai(sow_spec(data), max_tokens=2000)
    sec = parse_json_response(spec_json)

    # 2. Verify before rendering
    issues = verify_output('sow', sec, data)
    if issues:
        print(f'[SOW VERIFY] {issues}')

    # 3. Build fresh Word document — no template dependency
    doc = Document()
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.2)
        section.right_margin  = Inches(1.2)

    today = sec.get('effective_date', datetime.today().strftime('%-d %B %Y'))
    client = sec.get('client', data.get('clientName','CLIENT'))
    project = sec.get('project', data.get('projectName','PROJECT'))

    def heading(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after  = Pt(4)
        r = p.add_run(text)
        r.bold = True; r.font.size = Pt(11); r.font.color.rgb = C_PURPLE
        p.add_run().add_break()

    def body(text, size=10):
        if not text: return
        for line in str(text).strip().split('\n'):
            line = line.strip()
            if not line: continue
            if line.startswith(('•','-','*')):
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(line.lstrip('•-* ').strip()).font.size = Pt(size)
            else:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(3)
                p.add_run(line).font.size = Pt(size)

    def rule():
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        p.add_run('─' * 72).font.color.rgb = C_GREY

    def kv(label, value, size=10):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        lr = p.add_run(f'{label}:  ')
        lr.bold = True; lr.font.size = Pt(size); lr.font.color.rgb = C_PURPLE
        vr = p.add_run(str(value) if value else PH(label.lower()))
        vr.font.size = Pt(size); vr.font.color.rgb = C_DARK

    # Cover block
    tp = doc.add_paragraph()
    tp.add_run('CYPHR').font.size = Pt(22)
    tp.runs[0].bold = True; tp.runs[0].font.color.rgb = C_PURPLE

    sp = doc.add_paragraph()
    sr = sp.add_run('Statement of Work')
    sr.bold = True; sr.font.size = Pt(13); sr.font.color.rgb = C_DARK

    rule()
    kv('Effective date', today)
    kv('Client',         client)
    kv('Project',        project)
    kv('Prepared by',    f'Cyphr Studio — {CYPHR_EMAIL}')
    doc.add_paragraph()

    # Sections
    sections_map = [
        ('1. Project Summary',       sec.get('summary')),
        ('2. Objectives & Outputs',  sec.get('objectives')),
        ('3. Assumptions',           sec.get('assumptions')),
        ('4. Responsibilities',      sec.get('responsibilities')),
        ('5. Location',              sec.get('location', 'United Kingdom')),
        ('6. Project Milestones',    sec.get('milestones')),
        ('7. Risks',                 'Cyphr will address Severity 1 issues within 24 hours. Non-blocking defects may be deferred pending prioritisation.'),
    ]
    for title, content in sections_map:
        heading(title)
        body(content)

    # Commercial
    heading('8. Commercial')
    doc.add_paragraph()
    kv('Fee',            sec.get('fee', PH('fee — confirm with estimate')))
    kv('Payment terms',  sec.get('payment_terms', 'Invoiced at agreed milestones. Payment terms: 30 days from invoice date.'))
    doc.add_paragraph()

    # Signatures
    heading('9. Signatures')
    sig_table = doc.add_table(rows=2, cols=2)
    sig_table.style = 'Table Grid'
    headers = ['Signed for Cyphr Studio', f'Signed for {client}']
    for i, h in enumerate(headers):
        cell = sig_table.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.add_run(h).bold = True

    cyphr_cell = sig_table.rows[1].cells[0]
    client_cell = sig_table.rows[1].cells[1]
    for cell, lines in [
        (cyphr_cell, [CYPHR_DELIVERY_LEAD, CYPHR_DELIVERY_LEAD_ROLE, today]),
        (client_cell, [sec.get('client_contact', PH('client name')), PH('client role'), PH('date')]),
    ]:
        for line in lines:
            p = cell.add_paragraph()
            p.add_run(line).font.size = Pt(10)

    doc.add_paragraph()
    fp = doc.add_paragraph()
    fp.paragraph_format.space_before = Pt(20)
    fr = fp.add_run(f'Cyphr Studio  |  {CYPHR_EMAIL}  |  {CYPHR_SITE}  |  Confidential')
    fr.font.size = Pt(8); fr.font.color.rgb = C_GREY

    # Attach verify issues as a comment note at end if any
    if issues:
        doc.add_paragraph()
        note_p = doc.add_paragraph()
        nr = note_p.add_run('⚠ VERIFY BEFORE SENDING: ' + ' | '.join(issues))
        nr.font.size = Pt(8); nr.font.color.rgb = RGBColor(0xCC, 0x44, 0x00)

    doc.save(out)


def build_proposal_docx(data, out):
    # 1. Generate via AI
    spec_json = call_ai(proposal_spec(data), max_tokens=1500)
    sec = parse_json_response(spec_json)

    issues = verify_output('proposal', sec, data)
    if issues:
        print(f'[PROPOSAL VERIFY] {issues}')

    doc = Document()
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.25)
        section.right_margin  = Inches(1.25)

    client  = sec.get('client',  data.get('clientName','CLIENT'))
    project = sec.get('project', data.get('projectName','PROJECT'))
    today   = sec.get('date',    datetime.today().strftime('%-d %B %Y'))

    def heading(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after  = Pt(4)
        r = p.add_run(text)
        r.bold = True; r.font.size = Pt(11); r.font.color.rgb = C_PURPLE

    def body(text, size=10):
        if not text: return
        if isinstance(text, list):
            for item in text:
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(str(item)).font.size = Pt(size)
        else:
            for line in str(text).strip().split('\n'):
                line = line.strip()
                if not line: continue
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(3)
                p.add_run(line).font.size = Pt(size)

    def rule():
        p = doc.add_paragraph()
        p.add_run('─' * 72).font.color.rgb = C_GREY

    # Cover
    tp = doc.add_paragraph()
    tp.add_run('CYPHR').font.size = Pt(22)
    tp.runs[0].bold = True; tp.runs[0].font.color.rgb = C_PURPLE

    sp = doc.add_paragraph()
    sr = sp.add_run(f'PROPOSAL — {client.upper()}')
    sr.bold = True; sr.font.size = Pt(13); sr.font.color.rgb = C_DARK

    pp = doc.add_paragraph()
    pr = pp.add_run(project)
    pr.font.size = Pt(10); pr.font.color.rgb = C_GREY

    rule()

    def kv(label, value):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        lr = p.add_run(f'{label}:  ')
        lr.bold = True; lr.font.size = Pt(10); lr.font.color.rgb = C_PURPLE
        vr = p.add_run(str(value))
        vr.font.size = Pt(10); vr.font.color.rgb = C_DARK

    kv('Prepared for', client)
    kv('Project',      project)
    kv('Date',         today)
    kv('Investment',   sec.get('investment', PH('fee — confirm with estimate')))
    kv('Timeline',     sec.get('timeline',   PH('timeline — confirm at kick-off')))
    doc.add_paragraph()

    heading('Executive Summary')
    body(sec.get('executive_summary'))

    heading('The Opportunity')
    body(sec.get('opportunity'))

    heading('Our Approach')
    body(sec.get('approach'))

    heading('What We Deliver')
    body(sec.get('deliverables'))

    heading('Why Cyphr')
    body(sec.get('why_cyphr'))

    doc.add_paragraph()
    fp = doc.add_paragraph()
    fp.paragraph_format.space_before = Pt(24)
    fr = fp.add_run(f'Cyphr Studio  |  {CYPHR_EMAIL}  |  {CYPHR_SITE}  |  Confidential')
    fr.font.size = Pt(8); fr.font.color.rgb = C_GREY

    if issues:
        doc.add_paragraph()
        note_p = doc.add_paragraph()
        nr = note_p.add_run('⚠ VERIFY BEFORE SENDING: ' + ' | '.join(issues))
        nr.font.size = Pt(8); nr.font.color.rgb = RGBColor(0xCC, 0x44, 0x00)

    doc.save(out)


def build_estimate(data, out):
    # 1. Generate phase/role breakdown via AI
    spec_json = call_ai(estimate_spec(data), max_tokens=1500)
    sec = parse_json_response(spec_json)

    issues = verify_output('estimate', sec, data)
    if issues:
        print(f'[ESTIMATE VERIFY] {issues}')

    client  = sec.get('client',  data.get('clientName','CLIENT'))
    project = sec.get('project', data.get('projectName','PROJECT'))
    today   = datetime.today().strftime('%d/%m/%Y')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Estimate'

    # Styles
    def cell_style(cell, bold=False, bg=None, color='1F1F1F', size=10, align='left'):
        cell.font = Font(name='Arial', size=size, bold=bold,
                         color=color)
        if bg:
            cell.fill = PatternFill('solid', fgColor=bg)
        cell.alignment = Alignment(horizontal=align, vertical='center',
                                   wrap_text=True)

    def border_all(cell, color='D1D5DB'):
        s = Side(style='thin', color=color)
        cell.border = Border(left=s, right=s, top=s, bottom=s)

    # Header rows
    ws.merge_cells('A1:F1')
    ws['A1'] = f'Cyphr Studio — Cost Estimate'
    cell_style(ws['A1'], bold=True, size=14, color='5B4FD9')

    ws.merge_cells('A2:F2')
    ws['A2'] = f'{client} / {project}  |  {today}'
    cell_style(ws['A2'], size=10, color='6B7280')

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 18

    # Column headers row 4
    headers = ['Phase', 'Role', 'Day Rate (£)', 'Days', 'Cost (£)', 'Notes']
    col_widths = [22, 28, 14, 8, 12, 30]
    for i, (h, w) in enumerate(zip(headers, col_widths), 1):
        c = ws.cell(row=4, column=i, value=h)
        cell_style(c, bold=True, bg='5B4FD9', color='FFFFFF', align='center')
        border_all(c, '5B4FD9')
        ws.column_dimensions[get_column_letter(i)].width = w

    # Data rows
    row = 5
    grand_total = 0
    phases = sec.get('phases', [])

    for phase in phases:
        phase_name = phase.get('name', 'Phase')
        roles = phase.get('roles', [])
        phase_total = 0

        # Phase header row
        ws.merge_cells(f'A{row}:F{row}')
        c = ws.cell(row=row, column=1, value=phase_name)
        cell_style(c, bold=True, bg='E8E5FF', color='5B4FD9')
        ws.row_dimensions[row].height = 18
        row += 1

        for role_entry in roles:
            team_key = role_entry.get('team_key', '')
            days     = role_entry.get('days', 0)
            team     = CYPHR_TEAM.get(team_key)
            if not team:
                continue
            rate  = team['rate']
            cost  = rate * days
            phase_total += cost

            cells_data = [phase_name, team['name'], rate, days, cost, '']
            for col_i, val in enumerate(cells_data, 1):
                c = ws.cell(row=row, column=col_i, value=val)
                bg = 'F3F2FF' if (row % 2 == 0) else 'FFFFFF'
                cell_style(c, bg=bg, align='center' if col_i > 2 else 'left')
                border_all(c)
                if col_i in (3, 5):
                    c.number_format = '£#,##0'
            ws.row_dimensions[row].height = 16
            row += 1

        # Phase subtotal
        st_cell = ws.cell(row=row, column=4, value='Subtotal')
        cell_style(st_cell, bold=True, align='right')
        cost_cell = ws.cell(row=row, column=5, value=phase_total)
        cell_style(cost_cell, bold=True, bg='E8E5FF')
        cost_cell.number_format = '£#,##0'
        border_all(cost_cell, '5B4FD9')
        grand_total += phase_total
        row += 2

    # Grand total
    ws.merge_cells(f'A{row}:D{row}')
    tc = ws.cell(row=row, column=1, value='TOTAL')
    cell_style(tc, bold=True, size=11, bg='5B4FD9', color='FFFFFF', align='right')
    gc = ws.cell(row=row, column=5, value=grand_total)
    cell_style(gc, bold=True, size=11, bg='5B4FD9', color='FFFFFF', align='center')
    gc.number_format = '£#,##0'
    ws.row_dimensions[row].height = 22
    row += 2

    # Assumptions
    ws.cell(row=row, column=1, value='Assumptions').font = Font(name='Arial', bold=True, size=10, color='5B4FD9')
    row += 1
    for a in sec.get('assumptions', []):
        c = ws.cell(row=row, column=1, value=f'• {a}')
        c.font = Font(name='Arial', size=9, color='6B7280')
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

    row += 1
    ws.cell(row=row, column=1, value='Exclusions').font = Font(name='Arial', bold=True, size=10, color='5B4FD9')
    row += 1
    for e in sec.get('exclusions', []):
        c = ws.cell(row=row, column=1, value=f'• {e}')
        c.font = Font(name='Arial', size=9, color='6B7280')
        ws.merge_cells(f'A{row}:F{row}')
        row += 1

    # Rate card sheet
    rc = wb.create_sheet('Rate Card')
    rc_headers = ['Role', 'Day Rate (£)']
    for i, h in enumerate(rc_headers, 1):
        c = rc.cell(row=1, column=i, value=h)
        c.font = Font(name='Arial', bold=True, color='FFFFFF', size=10)
        c.fill = PatternFill('solid', fgColor='5B4FD9')
    for r_i, (key, val) in enumerate(CYPHR_TEAM.items(), 2):
        rc.cell(row=r_i, column=1, value=val['name']).font = Font(name='Arial', size=10)
        rate_c = rc.cell(row=r_i, column=2, value=val['rate'])
        rate_c.font = Font(name='Arial', size=10)
        rate_c.number_format = '£#,##0'
    rc.column_dimensions['A'].width = 30
    rc.column_dimensions['B'].width = 14

    if issues:
        note_sheet = wb.create_sheet('⚠ Verify')
        note_sheet['A1'] = 'Issues to check before sending:'
        note_sheet['A1'].font = Font(name='Arial', bold=True, color='CC4400', size=11)
        for i, issue in enumerate(issues, 2):
            note_sheet[f'A{i}'] = f'• {issue}'
            note_sheet[f'A{i}'].font = Font(name='Arial', size=10)

    wb.save(out)


def build_brief(data, out):
    """Brief stays as-is — it works."""
    doc = Document()
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.2)
        section.right_margin  = Inches(1.2)

    client     = data.get('clientName', 'CLIENT')
    project    = data.get('projectName', 'PROJECT')
    sector     = data.get('sector', '')
    timeline   = data.get('timeline', '')
    brief_text = data.get('briefOutput', '')
    requirements = data.get('requirements', '')
    bg_notes   = data.get('bgNotes', '')
    today      = datetime.today().strftime('%-d %B %Y')

    tp = doc.add_paragraph()
    r = tp.add_run('CYPHR')
    r.bold = True; r.font.size = Pt(22); r.font.color.rgb = C_PURPLE

    sp = doc.add_paragraph()
    r2 = sp.add_run(f'BRIEF — {client.upper()}')
    r2.bold = True; r2.font.size = Pt(11); r2.font.color.rgb = C_GREY

    if project:
        pp = doc.add_paragraph()
        r3 = pp.add_run(project)
        r3.font.size = Pt(10); r3.font.color.rgb = C_GREY

    doc.add_paragraph('─' * 68)

    def kv(label, value):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        lbl = p.add_run(f'{label}:  ')
        lbl.bold = True; lbl.font.color.rgb = C_PURPLE; lbl.font.size = Pt(10)
        val = p.add_run(str(value))
        val.font.size = Pt(10); val.font.color.rgb = C_DARK

    kv('Client',   client)
    kv('Project',  project)
    kv('Sector',   sector   or PH('sector'))
    kv('Budget',   f"£{int(float(str(data.get('budget','0')).replace('£','').replace(',',''))):,}" if data.get('budget') and str(data.get('budget')) not in ('0','') else PH('budget'))
    kv('Timeline', timeline or PH('timeline'))
    kv('Date',     today)
    doc.add_paragraph()

    if brief_text:
        for line in brief_text.strip().split('\n'):
            line = line.strip()
            if not line: doc.add_paragraph(); continue
            if line.startswith('##') or (line.startswith('**') and line.endswith('**')) or (line.isupper() and 3 < len(line) < 60):
                clean = line.lstrip('#').strip().strip('*')
                h = doc.add_paragraph()
                h.paragraph_format.space_before = Pt(14)
                hr = h.add_run(clean)
                hr.bold = True; hr.font.size = Pt(11); hr.font.color.rgb = C_PURPLE
            elif line.startswith(('• ','- ','* ')):
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(line[2:].strip()).font.size = Pt(10)
            else:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(4)
                p.add_run(line).font.size = Pt(10)
    elif requirements:
        h = doc.add_paragraph()
        h.add_run('REQUIREMENTS').bold = True
        for line in requirements.split('\n'):
            if line.strip():
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(line.strip().lstrip('•-* ')).font.size = Pt(10)

    doc.add_paragraph()
    fp = doc.add_paragraph()
    fp.paragraph_format.space_before = Pt(24)
    fr = fp.add_run(f'Prepared by Cyphr Studio  |  {CYPHR_EMAIL}  |  {today}')
    fr.font.size = Pt(8); fr.font.color.rgb = C_GREY
    doc.save(out)


def build_gantt(data, out):
    """Gantt stays as-is — it works. Just clear old data properly."""
    shutil.copy(f'{TEMPLATES_DIR}/gantt.xlsx', out)
    wb    = openpyxl.load_workbook(out)
    client  = data.get('clientName', 'CLIENT')
    project = data.get('projectName', 'PROJECT')
    label   = f'{client.upper()} — {project.upper()}'
    today = datetime.today()
    from datetime import timedelta
    days_ahead = (7 - today.weekday()) % 7 or 7
    week_start = today + timedelta(days=days_ahead)

    for ws in wb.worksheets:
        if 'decision' in ws.title.lower() or 'log' in ws.title.lower():
            for ri in range(2, ws.max_row + 1):
                for ci in range(1, ws.max_column + 1):
                    ws.cell(ri, ci).value = None
            continue
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    cell.value = (cell.value
                        .replace('HP CRUSH 3.0 — PROJECT GANTT', label)
                        .replace('HP CRUSH 3.0', client.upper())
                        .replace('HP Crush 3.0', client)
                        .replace('CLIENT', client)
                        .replace('PROJECT', project))
                    m = re.match(r'^W(\d+):', str(cell.value))
                    if m:
                        wn = int(m.group(1))
                        wdate = week_start + timedelta(weeks=wn-1)
                        cell.value = f'W{wn}: {wdate.strftime("%d %b")}'
        ws.title = f'{client} Gantt'
    wb.save(out)


# ═══════════════════════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'provider': os.environ.get('AI_PROVIDER','anthropic')})

@app.route('/generate', methods=['POST','OPTIONS'])
def generate():
    if request.method == 'OPTIONS':
        return '', 204

    body  = request.get_json()
    ftype = body.get('type')
    data  = body.get('projectData', {})
    slug  = re.sub(r'[^a-z0-9_]', '', (data.get('clientName') or 'project').lower().replace(' ','_'))

    ok, result = validate(ftype, data)
    if not ok:
        return jsonify(result), 400

    with tempfile.TemporaryDirectory() as tmp:
        try:
            if ftype == 'sow':
                out = f'{tmp}/{slug}_sow.docx'
                build_sow(data, out)
                return send_file(out, as_attachment=True, download_name=f'{slug}_sow.docx',
                               mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

            elif ftype in ('proposal', 'proposal-docx'):
                out = f'{tmp}/{slug}_proposal.docx'
                build_proposal_docx(data, out)
                return send_file(out, as_attachment=True, download_name=f'{slug}_proposal.docx',
                               mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

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

            elif ftype == 'brief':
                out = f'{tmp}/{slug}_brief.docx'
                build_brief(data, out)
                return send_file(out, as_attachment=True, download_name=f'{slug}_brief.docx',
                               mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

            elif ftype == 'brief-pdf':
                docx_out = f'{tmp}/{slug}_brief.docx'
                pdf_out  = f'{tmp}/{slug}_brief.pdf'
                build_brief(data, docx_out)
                subprocess.run(['libreoffice','--headless','--convert-to','pdf','--outdir',tmp,docx_out],
                               capture_output=True, timeout=30)
                if not os.path.exists(pdf_out):
                    return jsonify({'error': 'PDF conversion failed'}), 500
                return send_file(pdf_out, as_attachment=True, download_name=f'{slug}_brief.pdf',
                               mimetype='application/pdf')

            return jsonify({'error': 'unknown type'}), 400

        except Exception as e:
            print(f'[ERROR {ftype}] {e}')
            return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
