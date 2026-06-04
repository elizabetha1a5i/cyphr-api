from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import json, os, shutil, subprocess, re, tempfile
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree
import openpyxl

app = Flask(__name__)
CORS(app)

TEMPLATES_DIR = '/app/templates'
SCRIPTS_DIR   = '/app/scripts'


# ═══════════════════════════════════════════════════════════════════════════════
# CYPHR TEAM CONFIG
# ───────────────────────────────────────────────────────────────────────────────
# This is the ONLY place team names, roles and rates should come from.
# Never invent or hardcode team members anywhere else in this file.
# If someone joins or leaves, update it here.
# ═══════════════════════════════════════════════════════════════════════════════

CYPHR_TEAM = {
    # key: what appears in the estimate template's role column
    # name: display name — must match exactly what's in the estimate.xlsx template
    # rate: day rate in GBP
    # location: UK or Albania (for rate card display)
    'strategy_rob':   {'name': 'Strategy Lead (Rob)',        'rate': 950, 'location': 'UK'},
    'strategy_james': {'name': 'Strategy Lead (James)',      'rate': 950, 'location': 'UK'},
    'tech_lead':      {'name': 'Tech Lead (Redian)',         'rate': 700, 'location': 'UK'},
    'tech_lead_p2':   {'name': 'Tech Lead (Redian)',         'rate': 500, 'location': 'UK'},  # reduced phase 2 rate — matches template
    'producer':       {'name': 'Producer (Verity)',          'rate': 500, 'location': 'UK'},
    'qa':             {'name': 'QA',                         'rate': 450, 'location': 'UK'},
    'support':        {'name': 'Support',                    'rate': 500, 'location': 'UK'},
    'hosting':        {'name': 'Hosting & 3rd Party Costs',  'rate': 500, 'location': 'N/A'},
    'dev_mid':        {'name': 'Developer (Mid)',            'rate': 200, 'location': 'Albania'},
    'dev_senior':     {'name': 'Developer (Senior)',         'rate': 300, 'location': 'Albania'},
}

# SOW project lead — only Verity's name appears in Cyphr-signed SOWs
CYPHR_DELIVERY_LEAD = 'Verity Smout'
CYPHR_DELIVERY_LEAD_ROLE = 'Head of Production'


# ═══════════════════════════════════════════════════════════════════════════════
# PLACEHOLDERS
# ───────────────────────────────────────────────────────────────────────────────
# Used wherever data is missing. Visible and searchable in every document.
# Format is consistent so the user can Ctrl+F before sending anything.
# ═══════════════════════════════════════════════════════════════════════════════

def PH(reason):
    return f'[CONFIRM: {reason}]'

PLACEHOLDERS = {
    'budget':           PH('budget — not yet confirmed'),
    'timeline':         PH('timeline — not yet confirmed'),
    'client_contact':   PH('client project lead name and email'),
    'client_address':   PH('client registered address'),
    'fee_detail':       PH('fee breakdown — confirm with estimate'),
    'milestone_dates':  PH('milestone dates — confirm at kick-off'),
    'sector':           PH('sector — verify extraction is correct'),
    'summary':          PH('project summary — re-run context extraction'),
}


# ═══════════════════════════════════════════════════════════════════════════════
# VALIDATION
# ───────────────────────────────────────────────────────────────────────────────
# Hard required fields per document type.
# Missing = friendly HTTP 400 asking the user to provide the information.
# Wrong/hallucinated data = checked by cross-referencing against what was sent.
# ═══════════════════════════════════════════════════════════════════════════════

REQUIRED_FIELDS = {
    'brief':    ['clientName', 'projectName'],
    'sow':      ['clientName', 'projectName', 'sowOutput'],
    'proposal': ['clientName', 'projectName'],
    'estimate': ['clientName', 'projectName'],
    'gantt':    ['clientName', 'projectName'],
}

# Minimum character length for AI-generated fields.
# Below this = extraction likely failed.
MIN_LENGTH = {
    'sowOutput':      100,
    'briefOutput':    80,
    'proposalOutput': 100,
}

# Human-readable explanations for each required field — shown to the user on 400
FIELD_EXPLANATIONS = {
    'clientName':    'the client name, so documents are addressed to the right organisation',
    'projectName':   'the project name, so every document refers to the right engagement',
    'sowOutput':     'the SOW content from the AI generation stage — please go back and run that step first',
    'briefOutput':   'the brief content from the AI generation stage — please go back and run that step first',
    'proposalOutput':'the proposal content from the AI generation stage — please go back and run that step first',
}


def validate(ftype, data):
    """
    Check required fields and AI output quality before generating.

    Returns:
        (True, warnings_list)   — ok to proceed, warnings is [] or has soft issues
        (False, error_payload)  — missing required data, return as HTTP 400
    """
    missing_explanations = []
    warnings = []

    # 1. Hard required fields
    for field in REQUIRED_FIELDS.get(ftype, []):
        val = str(data.get(field, '')).strip()
        if not val:
            explanation = FIELD_EXPLANATIONS.get(field, field)
            missing_explanations.append(explanation)

    # 2. AI output minimum length
    for field, min_len in MIN_LENGTH.items():
        val = str(data.get(field, '')).strip()
        if val and len(val) < min_len:
            missing_explanations.append(
                f'the {field.replace("Output","")} content looks incomplete '
                f'(only {len(val)} characters) — please re-run the generation step'
            )

    if missing_explanations:
        return False, {
            'error': 'A few things are needed before this document can be generated.',
            'needed': missing_explanations,
            'message': (
                'To generate this document Cyphr Flow needs: '
                + '; and '.join(missing_explanations) + '. '
                'Once these are filled in, come back and try again.'
            )
        }

    # 3. Soft warnings — these use placeholders, not hard fail
    if not str(data.get('budget', '')).strip() or str(data.get('budget', '')).strip() in ('0', '£0'):
        warnings.append('budget not set — document will show a placeholder')
    if not str(data.get('timeline', '')).strip():
        warnings.append('timeline not set — document will show a placeholder')

    # 4. Hallucination / data integrity check
    #    Cross-reference the AI outputs against the source data the user sent.
    #    If team member names appear in AI output that aren't in CYPHR_TEAM,
    #    flag it — the AI has likely invented someone.
    hallucination_warnings = check_for_invented_content(data)
    warnings.extend(hallucination_warnings)

    return True, warnings


def check_for_invented_content(data):
    """
    Check AI-generated text fields for content that wasn't in the source data
    and isn't in the known Cyphr config.

    Specifically:
    - Person names that aren't in CYPHR_TEAM or the source documents
    - Roles that aren't in CYPHR_TEAM
    - Budget/fee figures that don't match what the user sent

    Returns a list of warning strings. Empty = clean.
    """
    warnings = []

    # Build the set of names/roles we know are legitimate
    known_names = {v['name'].lower() for v in CYPHR_TEAM.values()}
    known_names.add(CYPHR_DELIVERY_LEAD.lower())

    # Also treat anything in the user's source documents as known
    source_text = ' '.join(filter(None, [
        str(data.get('briefOutput', '')),
        str(data.get('requirements', '')),
        str(data.get('bgNotes', '')),
        str(data.get('clientName', '')),
        str(data.get('projectName', '')),
    ])).lower()

    # Check AI-generated outputs for person names not in known set or source text
    ai_outputs = ' '.join(filter(None, [
        str(data.get('sowOutput', '')),
        str(data.get('proposalOutput', '')),
        str(data.get('estimateOutput', '')),
    ]))

    # Look for patterns like "Name (Role)" or capitalised proper name pairs
    # that might indicate an invented person
    # Only flag two-word proper names that look like people:
    # both words 3+ chars, neither word is a common project/doc/legal term.
    NOT_NAMES = {
        'project', 'phase', 'support', 'retail', 'digital', 'experience', 'summary',
        'overview', 'assumptions', 'exclusions', 'objectives', 'milestones', 'commercial',
        'invoice', 'submission', 'instructions', 'management', 'reporting', 'responsibilities',
        'risks', 'conditions', 'agreement', 'contract', 'services', 'information', 'protection',
        'rights', 'property', 'intellectual', 'confidential', 'applicable', 'termination',
        'liability', 'indemnity', 'circumstances', 'communications', 'jurisdiction', 'severance',
        'base', 'main', 'consulting', 'corporation', 'internet', 'social', 'media', 'security',
        'breach', 'neither', 'party', 'effective', 'date', 'set', 'off', 'these', 'the',
        'new', 'old', 'east', 'west', 'north', 'south', 'united', 'kingdom', 'limited',
        'industries', 'camburgh', 'dover', 'accounts', 'payable', 'samsung', 'blue', 'square',
        'fee', 'total', 'silver', 'bronze', 'gold', 'tour', 'lab', 'app', 'store', 'flip',
        'fold', 'galaxy', 'build', 'launch', 'sprint', 'design', 'discovery', 'planning',
        'change', 'request', 'log', 'risk', 'output', 'input', 'phase', 'work',
    }
    name_pattern = re.compile(r'\b([A-Z][a-z]{2,} [A-Z][a-z]{2,})\b')
    found_names = set(name_pattern.findall(ai_outputs))

    for name in found_names:
        words = name.lower().split()
        # Skip if either word is a known non-name term
        if any(w in NOT_NAMES for w in words):
            continue
        if name.lower() in known_names:
            continue
        if name.lower() in source_text:
            continue
        warnings.append(
            f'"{name}" appears in the AI output but wasn\'t in your source documents '
            f'or Cyphr\'s team config — check this isn\'t an invented person or contact'
        )

    # Check that any fee/budget figures in the AI output aren't wildly different
    # from what the user sent (catches AI inventing a budget)
    raw_budget = str(data.get('budget', '')).replace('£', '').replace(',', '').strip()
    if raw_budget and raw_budget not in ('0', ''):
        try:
            user_budget = int(float(raw_budget))
            # Find any £ amounts in AI outputs
            fee_pattern = re.compile(r'£([\d,]+)')
            ai_fees = [int(m.replace(',', '')) for m in fee_pattern.findall(ai_outputs)]
            for fee in ai_fees:
                # Only flag figures that EXCEED the total budget by more than 20%
                # (line items in a fee breakdown are expected to be smaller than the total)
                if user_budget > 0 and fee > user_budget * 1.2 and fee > 1000:
                    warnings.append(
                        f'AI output contains £{fee:,} which exceeds the project budget of '
                        f'£{user_budget:,} — check this figure is intentional'
                    )
        except (ValueError, ZeroDivisionError):
            pass

    return warnings


# ═══════════════════════════════════════════════════════════════════════════════
# SAFE FIELD GETTERS
# ═══════════════════════════════════════════════════════════════════════════════

def fmt(budget):
    try: return f"{int(float(str(budget).replace('£','').replace(',',''))):,}"
    except: return str(budget)

def weeks(timeline):
    m = re.search(r'(\d+)', str(timeline))
    return int(m.group(1)) if m else 10

def safe_budget(data):
    raw = str(data.get('budget', '')).strip().replace('£', '').replace(',', '')
    if not raw or raw == '0':
        return PLACEHOLDERS['budget']
    try:
        val = int(float(raw))
        return f'£{val:,}' if val > 0 else PLACEHOLDERS['budget']
    except:
        return str(data.get('budget', '')) or PLACEHOLDERS['budget']

def safe_summary(data, sec):
    s = sec.get('summary', '').strip()
    return s if len(s) > 20 else PLACEHOLDERS['summary']

def safe_fee(data, sec):
    fee = sec.get('fee', '').strip()
    if not fee:
        return PLACEHOLDERS['fee_detail']
    if re.search(r'£\s*0\b', fee) and len(fee) < 20:
        return PLACEHOLDERS['fee_detail']
    # Reject if it looks like research methodology rather than a fee summary.
    # A valid fee section should mention a price, total, or payment term.
    # If it talks about recruitment, participants, or interviews it's the wrong section.
    bad_signals = ['recruit', 'participant', 'interview', 'n=', 'persona', 'deliver']
    fee_signals = ['£', 'total', 'invoice', 'payment', 'fixed price', 'fee', 'cost']
    fee_lower = fee.lower()
    has_fee_signal = any(s in fee_lower for s in fee_signals)
    has_bad_signal = any(s in fee_lower for s in bad_signals)
    if has_bad_signal and not has_fee_signal:
        return PLACEHOLDERS['fee_detail']
    return fee

def safe_milestones(data, sec):
    ms = sec.get('milestones', '').strip()
    return ms if (ms and len(ms.splitlines()) >= 2) else PLACEHOLDERS['milestone_dates']


# ═══════════════════════════════════════════════════════════════════════════════
# ROUTE
# ═══════════════════════════════════════════════════════════════════════════════

@app.route('/health')
def health():
    return jsonify({'status': 'ok'})

@app.route('/generate', methods=['POST', 'OPTIONS'])
def generate():
    if request.method == 'OPTIONS':
        return '', 204

    body  = request.get_json()
    ftype = body.get('type')
    data  = body.get('projectData', {})
    slug  = re.sub(r'[^a-z0-9_]', '', (data.get('clientName') or 'project').lower().replace(' ', '_'))

    ok, result = validate(ftype, data)
    if not ok:
        # Friendly 400 — tells the user exactly what's needed and why
        return jsonify(result), 400

    # Log any warnings (will appear as placeholders in the document)
    if result:
        print(f'[WARNINGS {ftype}/{slug}] {result}')

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
        elif ftype == 'brief':
            out = f'{tmp}/{slug}_brief.docx'
            build_brief(data, out)
            return send_file(out, as_attachment=True, download_name=f'{slug}_brief.docx',
                           mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        elif ftype == 'brief-pdf':
            docx_out = f'{tmp}/{slug}_brief.docx'
            pdf_out  = f'{tmp}/{slug}_brief.pdf'
            build_brief(data, docx_out)
            subprocess.run(
                ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, docx_out],
                capture_output=True, timeout=30
            )
            if not os.path.exists(pdf_out):
                return jsonify({'error': 'PDF conversion failed — LibreOffice not available'}), 500
            return send_file(pdf_out, as_attachment=True, download_name=f'{slug}_brief.pdf',
                           mimetype='application/pdf')

        return jsonify({'error': 'unknown type'}), 400


# ═══════════════════════════════════════════════════════════════════════════════
# BRIEF
# ═══════════════════════════════════════════════════════════════════════════════

def build_brief(data, out):
    doc = Document()
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.2)
        section.right_margin  = Inches(1.2)

    CYPHR_GREEN = RGBColor(0x4A, 0x7C, 0x59)
    DARK        = RGBColor(0x1A, 0x1A, 0x1A)
    MID         = RGBColor(0x55, 0x55, 0x55)

    client     = data.get('clientName', 'CLIENT')
    project    = data.get('projectName', 'PROJECT')
    sector     = data.get('sector', '')
    timeline   = data.get('timeline', '')
    brief_text = data.get('briefOutput', '')
    requirements = data.get('requirements', '')
    bg_notes   = data.get('bgNotes', '')
    today      = datetime.today().strftime('%-d %B %Y')

    title_p = doc.add_paragraph()
    r = title_p.add_run('CYPHR')
    r.bold = True; r.font.size = Pt(22); r.font.color.rgb = CYPHR_GREEN

    sub_p = doc.add_paragraph()
    r2 = sub_p.add_run(f'BRIEF — {client.upper()}')
    r2.bold = True; r2.font.size = Pt(11); r2.font.color.rgb = MID

    if project:
        proj_p = doc.add_paragraph()
        r3 = proj_p.add_run(project)
        r3.font.size = Pt(10); r3.font.color.rgb = MID

    doc.add_paragraph('─' * 68)

    def meta(label, value):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        lbl = p.add_run(f'{label}:  ')
        lbl.bold = True; lbl.font.color.rgb = CYPHR_GREEN; lbl.font.size = Pt(10)
        val = p.add_run(str(value))
        val.font.size = Pt(10); val.font.color.rgb = DARK

    meta('Client',   client)
    meta('Project',  project)
    meta('Sector',   sector   if sector   else PLACEHOLDERS['sector'])
    meta('Budget',   safe_budget(data))
    meta('Timeline', timeline if timeline else PLACEHOLDERS['timeline'])
    meta('Date',     today)
    doc.add_paragraph()

    if brief_text:
        for line in brief_text.strip().split('\n'):
            line = line.strip()
            if not line:
                doc.add_paragraph(); continue
            if (line.startswith('##') or
                (line.startswith('**') and line.endswith('**')) or
                (line.isupper() and 3 < len(line) < 60)):
                clean = line.lstrip('#').strip().strip('*')
                h = doc.add_paragraph()
                h.paragraph_format.space_before = Pt(14)
                h.paragraph_format.space_after  = Pt(4)
                hr = h.add_run(clean)
                hr.bold = True; hr.font.size = Pt(11); hr.font.color.rgb = CYPHR_GREEN
            elif line.startswith(('• ', '- ', '* ', '· ')):
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(line[2:].strip()).font.size = Pt(10)
            else:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(4)
                p.add_run(line).font.size = Pt(10)
    else:
        if requirements:
            h = doc.add_paragraph()
            h.add_run('REQUIREMENTS').bold = True
            for line in requirements.split('\n'):
                if line.strip():
                    p = doc.add_paragraph(style='List Bullet')
                    p.add_run(line.strip().lstrip('•-* ')).font.size = Pt(10)
        if bg_notes:
            h2 = doc.add_paragraph()
            h2.add_run('BACKGROUND & CONSIDERATIONS').bold = True
            doc.add_paragraph().add_run(bg_notes).font.size = Pt(10)

    doc.add_paragraph()
    footer_p = doc.add_paragraph()
    footer_p.paragraph_format.space_before = Pt(24)
    fr = footer_p.add_run(f'Prepared by Cyphr Studio  |  elizabeth@cyphr.studio  |  {today}')
    fr.font.size = Pt(8); fr.font.color.rgb = MID
    doc.save(out)


# ═══════════════════════════════════════════════════════════════════════════════
# SOW HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def replace_in_doc(doc, old, new):
    def _replace(para):
        for run in para.runs:
            if old in run.text:
                run.text = run.text.replace(old, new)
    for para in doc.paragraphs:
        _replace(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _replace(para)

def remove_hyperlink_runs(doc, fragment):
    WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    HL  = f'{{{WNS}}}hyperlink'
    T   = f'{{{WNS}}}t'
    def _clean(para):
        for child in list(para._p):
            if child.tag == HL:
                if fragment.lower() in ''.join(t.text or '' for t in child.iter(T)).lower():
                    para._p.remove(child)
    for para in doc.paragraphs:
        _clean(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _clean(para)

def set_cell(table, row_idx, text):
    """Write text to first cell of a row AND clear all other cells in that row."""
    if row_idx >= len(table.rows) or text is None:
        return
    WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    row = table.rows[row_idx]

    # Clear ALL cells in the row — old milestone dates live in cols 1 and 2
    for cell in row.cells:
        tc = cell._tc
        for p_el in tc.findall(f'{{{WNS}}}p'):
            tc.remove(p_el)
        # Add a blank paragraph so the cell isn't empty (required by OOXML)
        blank = OxmlElement('w:p')
        tc.append(blank)

    # Write the new text into the first cell only
    tc0 = row.cells[0]._tc
    # Remove the blank we just added
    for p_el in tc0.findall(f'{{{WNS}}}p'):
        tc0.remove(p_el)
    for line in str(text).split('\n'):
        p = OxmlElement('w:p')
        r = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.text = line
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(t); p.append(r); tc0.append(p)

def parse_sow(text):
    """Parse AI SOW output into sections. Returns only what's actually there."""
    s = {}
    if not text:
        return s
    cur, buf = None, []
    keys = {
        '1.1': 'summary',      'project summary': 'summary',
        '1.2': 'objectives',   'objectives': 'objectives',
        '1.3': 'assumptions',  'assumptions': 'assumptions',
        'responsibilities':    'responsibilities',
        '4.1': 'milestones',   'milestones': 'milestones',
        '5.1': 'fee',          'fee summary': 'fee', '3.1': 'fee',
    }
    for line in text.split('\n'):
        clean = re.sub(r'^#{1,6}\s*', '', line).strip().strip('*').strip('_')
        lo = clean.lower()
        matched = False
        for k, sec in keys.items():
            if lo.startswith(k):
                if cur and buf: s[cur] = '\n'.join(buf).strip()
                cur, buf, matched = sec, [], True
                break
        if not matched and cur and clean:
            buf.append(clean)
    if cur and buf:
        s[cur] = '\n'.join(buf).strip()
    return s


# ═══════════════════════════════════════════════════════════════════════════════
# SOW
# ═══════════════════════════════════════════════════════════════════════════════

def build_sow(data, out):
    doc     = Document(f'{TEMPLATES_DIR}/sow.docx')
    client  = data.get('clientName', 'CLIENT')
    project = data.get('projectName', 'PROJECT')
    timeline = data.get('timeline', '')
    sow     = data.get('sowOutput', '')
    today   = datetime.today().strftime('%-d %B %Y')

    # Only replace Blue Square-specific strings.
    # Do NOT replace 'Samsung' — it's the end client in the T&C boilerplate
    # and replacing it creates nonsense like "Samsung's client, Samsung".
    for old, new in [
        ('Blue Square Marketing Limited', client),
        ('Blue Square', client),
        ('AccountsPayable@bluesquare.uk.com', f'accounts@{client.lower().replace(" ","")}.com'),
        ('bluesquare.uk.com', f'{client.lower().replace(" ","")}.com'),
        ('Tate House, Watermark Way, Hertford SG13 7TZ', PLACEHOLDERS['client_address']),
        ('20th March 2026', today),
        # FIX: also replace the 28th April signature date and the footer creation date
        ('28th April 2026', today),
        ('27/01/2026', today),
        ('Samsung 2026 Q2 Contact Centre Roadshow', project),
        ('roadshow web app and photo booth experience', project),
        ('roadshow', 'project'),
        ('Roadshow', 'Project'),
    ]:
        replace_in_doc(doc, old, new)

    remove_hyperlink_runs(doc, 'Charlotte')
    remove_hyperlink_runs(doc, 'charlotte.cavanagh')
    replace_in_doc(doc, 'Charlotte Cavanagh', PLACEHOLDERS['client_contact'])

    t   = doc.tables[0]
    sec = parse_sow(sow)

    rmap = {
        3:  safe_summary(data, sec),
        5:  sec.get('objectives',
                f'Deliver {project} for {client} on time and within agreed budget.'),
        7:  sec.get('assumptions',
                f'Client to provide all required content and access within agreed timelines.\n'
                f'All third-party integrations and API access to be arranged by {client} prior to kick-off.'),
        9:  sec.get('responsibilities',
                f'Cyphr: design, build and delivery.\n'
                f'{client}: content provision, stakeholder sign-off and UAT feedback.'),
        11: 'United Kingdom',
        15: (f'Cyphr will meet with the {client} team regularly to discuss requirements and progress. '
             f'Weekly status updates will be provided throughout the project.'),
        # Delivery lead comes from config, not invented
        17: f'Cyphr: {CYPHR_DELIVERY_LEAD}\n{client}: {PLACEHOLDERS["client_contact"]}',
        19: ('Cyphr will provide regular project updates during delivery. '
             'A shared project tracker will be maintained throughout.'),
        21: safe_milestones(data, sec),
        23: 'Cyphr will address critical issues within 24 hours. All bugs tracked within agreed SLAs.',
        26: safe_fee(data, sec),
        28: 'Invoiced at project milestones. Payment terms: 30 days from invoice date.',
        30: f'Invoice to: Accounts Payable | {client}',
        32: 'Change requests require written approval before work commences.',
        # Row 36 = signature date line — overwrite with today's date
        36: today,
    }
    for ri, txt in rmap.items():
        set_cell(t, ri, txt)

    doc.save(out)


# ═══════════════════════════════════════════════════════════════════════════════
# PROPOSAL
# ═══════════════════════════════════════════════════════════════════════════════

def build_proposal(data, out, tmp):
    client   = data.get('clientName', 'CLIENT').upper()
    project  = data.get('projectName', 'PROJECT')
    timeline = data.get('timeline', '') or PLACEHOLDERS['timeline']
    proposal = data.get('proposalOutput', '')
    brief    = data.get('briefOutput', '')

    work = f'{tmp}/pwork/'
    result = subprocess.run(
        ['python3', f'{SCRIPTS_DIR}/unpack.py', f'{TEMPLATES_DIR}/proposal.pptx', work],
        capture_output=True, text=True, cwd=SCRIPTS_DIR
    )
    if result.returncode != 0:
        raise Exception(f'Unpack failed: {result.stderr}')

    slides = f'{work}ppt/slides/'
    paras  = [p.strip() for p in (proposal or brief or '').split('\n\n') if p.strip()]
    exec_sum = paras[0][:400] if paras else PH('executive summary — re-run proposal stage')
    the_ask  = paras[1][:300] if len(paras) > 1 else PH('project ask — re-run proposal stage')

    replacements = {
        'slide1.xml': [('CLIENT', client), ('PROJECT NAME', project.upper()), ('Cost Estimate ', 'Commercial Proposal ')],
        'slide2.xml': [('This proposal outlines….', exec_sum)],
        'slide3.xml': [('Activity Overview….', the_ask)],
        'slide8.xml': [
            ('Core Roadshow Experience Web App Design &amp; Build', project),
            ('Core Roadshow Experience Web App Design & Build', project),
            ('£25,314', safe_budget(data)),
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



# ═══════════════════════════════════════════════════════════════════════════════
# PROPOSAL (.docx) — Word version alongside the .pptx
# ═══════════════════════════════════════════════════════════════════════════════

def build_proposal_docx(data, out):
    """Generate a Word proposal document from the AI-produced proposalOutput."""
    from docx.oxml import OxmlElement
    doc = Document()
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.25)
        section.right_margin  = Inches(1.25)

    CYPHR_GREEN = RGBColor(0x4A, 0x7C, 0x59)
    DARK        = RGBColor(0x1A, 0x1A, 0x1A)
    MID         = RGBColor(0x55, 0x55, 0x55)

    client   = data.get('clientName', 'CLIENT')
    project  = data.get('projectName', 'PROJECT')
    proposal = data.get('proposalOutput', '')
    brief    = data.get('briefOutput', '')
    today    = datetime.today().strftime('%-d %B %Y')

    # Header
    title_p = doc.add_paragraph()
    title_p.add_run('CYPHR').font.size = Pt(22)
    title_p.runs[0].bold = True
    title_p.runs[0].font.color.rgb = CYPHR_GREEN

    sub_p = doc.add_paragraph()
    r = sub_p.add_run(f'PROPOSAL — {client.upper()}')
    r.bold = True; r.font.size = Pt(11); r.font.color.rgb = MID

    proj_p = doc.add_paragraph()
    r2 = proj_p.add_run(project)
    r2.font.size = Pt(10); r2.font.color.rgb = MID

    doc.add_paragraph('─' * 68)

    def meta(label, value):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        lbl = p.add_run(f'{label}:  ')
        lbl.bold = True; lbl.font.color.rgb = CYPHR_GREEN; lbl.font.size = Pt(10)
        val = p.add_run(str(value))
        val.font.size = Pt(10); val.font.color.rgb = DARK

    meta('Prepared for', client)
    meta('Project',      project)
    meta('Budget',       safe_budget(data))
    meta('Timeline',     data.get('timeline', '') or PLACEHOLDERS['timeline'])
    meta('Date',         today)
    doc.add_paragraph()

    # Body — use proposalOutput if available, fall back to briefOutput
    body_text = proposal or brief or ''
    if body_text:
        for line in body_text.strip().split('\n'):
            line = line.strip()
            if not line:
                doc.add_paragraph(); continue
            if (line.startswith('##') or
                (line.startswith('**') and line.endswith('**')) or
                (line.isupper() and 3 < len(line) < 60)):
                clean = line.lstrip('#').strip().strip('*')
                h = doc.add_paragraph()
                h.paragraph_format.space_before = Pt(14)
                h.paragraph_format.space_after  = Pt(4)
                hr = h.add_run(clean)
                hr.bold = True; hr.font.size = Pt(11); hr.font.color.rgb = CYPHR_GREEN
            elif line.startswith(('• ', '- ', '* ', '· ')):
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(line[2:].strip()).font.size = Pt(10)
            else:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(4)
                p.add_run(line).font.size = Pt(10)
    else:
        doc.add_paragraph().add_run(
            PLACEHOLDERS['summary']
        ).font.size = Pt(10)

    doc.add_paragraph()
    footer_p = doc.add_paragraph()
    footer_p.paragraph_format.space_before = Pt(24)
    fr = footer_p.add_run(
        f'Prepared by Cyphr Studio  |  elizabeth@cyphr.studio  |  {today}\n'
        f'Confidential. Not for distribution.'
    )
    fr.font.size = Pt(8); fr.font.color.rgb = MID
    doc.save(out)

# ═══════════════════════════════════════════════════════════════════════════════
# ESTIMATE
# ═══════════════════════════════════════════════════════════════════════════════

def build_estimate(data, out):
    shutil.copy(f'{TEMPLATES_DIR}/estimate.xlsx', out)
    wb = openpyxl.load_workbook(out)
    ws = wb['TEMPLATE']
    client  = data.get('clientName', 'CLIENT')
    project = data.get('projectName', 'PROJECT')
    today   = datetime.today().strftime('%d/%m/%Y')

    # Update header strings
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                if 'Cyphr Cost Estimate' in cell.value:
                    cell.value = f'Cyphr Cost Estimate — {client} / {project}'
                cell.value = (cell.value
                    .replace('CLIENT_NAME', client).replace('PROJECT_NAME', project)
                    .replace('CLIENT NAME', client).replace('PROJECT NAME', project))

    ws['G1'] = f'{client} / {project}'
    ws['G2'] = f'Date: {today}'

    # Write rate/days from CYPHR_TEAM config — not hardcoded names, not invented roles.
    # Row mapping matches the estimate.xlsx template structure exactly.
    # If you change the template, update these row numbers to match.
    team = CYPHR_TEAM
    # Row mapping verified against the actual estimate.xlsx template
    phase1_rows = {
        6:  ('strategy_rob',   0.5),   # Strategy Lead (Rob)     rate 950
        7:  ('strategy_james', 0.5),   # Strategy Lead (James)   rate 950
        8:  ('tech_lead',      0),     # Tech Lead (Redian)      rate 700
        9:  ('producer',       0.5),   # Producer (Verity)       rate 500
    }
    phase2_rows = {
        22: ('strategy_rob',   0),     # Strategy Lead (Rob)     rate 950
        23: ('strategy_james', 1),     # Strategy Lead (James)   rate 950
        24: ('producer',       2),     # Producer (Verity)       rate 500
        25: ('tech_lead_p2',   10),    # Tech Lead (Redian)      rate 500 (phase 2 rate — matches template)
        26: ('dev_mid',        0),     # Developer (Mid)         rate 200
        27: ('qa',             1),     # QA                      rate 450
        28: ('support',        1),     # Support                 rate 500
        29: ('hosting',        1),     # Hosting & 3rd Party     rate 500
    }
    for row_num, (team_key, default_days) in {**phase1_rows, **phase2_rows}.items():
        if team_key in team:
            rate = team[team_key]['rate']
            ws[f'C{row_num}'] = rate
            ws[f'D{row_num}'] = default_days
            # Preserve the template formula pattern (=Cx*Dx) rather than inventing new ones
            ws[f'E{row_num}'] = f'=C{row_num}*D{row_num}'
            ws[f'F{row_num}'] = f'=E{row_num}/(1-$O$5)'

    # Clear the Timings + Deliverables sheet — it contains old project data
    # from whatever project was used to build the template
    for sheet_name in wb.sheetnames:
        if 'timing' in sheet_name.lower() or 'deliverable' in sheet_name.lower():
            ws_clear = wb[sheet_name]
            for row in ws_clear.iter_rows():
                for cell in row:
                    cell.value = None

    wb.save(out)


# ═══════════════════════════════════════════════════════════════════════════════
# GANTT
# ═══════════════════════════════════════════════════════════════════════════════

def build_gantt(data, out):
    shutil.copy(f'{TEMPLATES_DIR}/gantt.xlsx', out)
    wb    = openpyxl.load_workbook(out)
    client  = data.get('clientName', 'CLIENT')
    project = data.get('projectName', 'PROJECT')
    label   = f'{client.upper()} — {project.upper()}'

    today = datetime.today()
    days_ahead = (7 - today.weekday()) % 7 or 7
    week_start = today + timedelta(days=days_ahead)

    for ws in wb.worksheets:
        # Clear Decision Log — never leave previous project data in it
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
                        .replace('SAMSUNG — SAMSUNG EXPERIENTIAL MARKETING CAMPAIGN', label)
                        .replace('HP CRUSH 3.0', client.upper())
                        .replace('HP Crush 3.0', client)
                        .replace('CLIENT', client)
                        .replace('PROJECT', project)
                    )
                    m = re.match(r'^W(\d+):', str(cell.value))
                    if m:
                        wn = int(m.group(1))
                        wdate = week_start + timedelta(weeks=wn - 1)
                        cell.value = f'W{wn}: {wdate.strftime("%d %b")}'

        ws.title = f'{client} Gantt'

    wb.save(out)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
