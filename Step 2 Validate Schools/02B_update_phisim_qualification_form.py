r"""Update the saved PHISIM University Qualification Form for each annual cycle.

Install once:
    py -m pip install openpyxl python-docx

From Step 2 Validate Schools:
    py .\02B_update_phisim_qualification_form.py --dry-run
    py .\02B_update_phisim_qualification_form.py
    py .\02B_update_phisim_qualification_form.py --serve

Selects the highest year in YYYY PHISIM Scholarship Process Dates.docx/.xlsx.
Reads the university qualification deadline, preserves PHISIM branding, removes
Academic Partners, and updates the three academic years, review, print, and
downloaded form. The academic year ending in the due-date year comes first.

The script updates local files. With --serve it serves the form on this computer
and saves submitted responses to the matching Member-Organization row of Current
evaluation checklist. Open the displayed HTTP address, not index.html directly.
The workbook must be closed in Excel. The public hosted site cannot reach this
local workbook. Before changes it backs up the form folder and before submissions
it backs up the workbook. It does not publish the hosted site or send email.
"""

import argparse
import base64
from datetime import date, datetime, timedelta
import json
import mimetypes
from pathlib import Path
import re
import shutil
import tempfile
import os
import secrets
import threading
import ctypes
from urllib.request import Request, urlopen
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlsplit
from zipfile import ZipFile


TITLE = 'PHISIM University Qualification Form'
WORKBOOK_WRITE_LOCK = threading.RLock()
REVIEW_DIALOG_CSS = '''
/* PHISIM visible review actions v1 */
@media screen {
 dialog[open]{display:flex;flex-direction:column;height:min(90vh,900px);overflow:hidden}
 .dialog-header{flex:0 0 auto;margin-bottom:12px}
 #review-content{flex:1 1 auto;min-height:70px;overflow-y:auto;padding-right:10px}
 dialog>.hint,dialog>#submission-status,dialog>#download-status{flex:0 0 auto}
 .dialog-actions{flex:0 0 auto;margin-top:12px}
 #submit-to-checklist{flex-basis:100%;width:100%}
 button:disabled{cursor:default;opacity:.7}
}
'''
SCHEDULE_FOLDER = 'Step 1  Generate Scholarship Dates and Request for Nominations'
VERSION = re.compile(r'^(\d{4})\s+PHISIM Scholarship Process Dates\.(docx|xlsx)$', re.I)
DATE_PATTERN = re.compile(r'\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})\b', re.I)

PREFILL_JS = r'''// PHISIM personalized form v1
let checklistPrefillError='';
const initialPrefillHash=new URLSearchParams(location.hash.slice(1)).get('prefill');
window.addEventListener('hashchange',()=>{const next=new URLSearchParams(location.hash.slice(1)).get('prefill');if(next&&next!==initialPrefillHash)location.reload();});
function readChecklistPrefill(){
 const encoded=new URLSearchParams(location.hash.slice(1)).get('prefill');
 if(!encoded)return null;
 try{
  if(encoded.length>24000||!/^[A-Za-z0-9_-]+$/.test(encoded))throw new Error();
  const raw=encoded.replace(/-/g,'+').replace(/_/g,'/');
  const decoded=new TextDecoder('utf-8',{fatal:true}).decode(Uint8Array.from(atob(raw+'='.repeat((4-raw.length%4)%4)),c=>c.charCodeAt(0)));
  const answers=JSON.parse(decoded);
  if(!answers||typeof answers!=='object'||Array.isArray(answers)||answers.v!==1||answers.year!==qualificationDueYear||typeof answers.university!=='string'||!answers.university.trim()||answers.university.length>200)throw new Error();
  const limits={university:200,school:200,faculty:200,email:254,programs:3000,accreditation:2000};
  const safe={year:String(answers.year)};
  for(const [field,max]of Object.entries(limits))if(typeof answers[field]==='string'&&answers[field].length<=max)safe[field]=answers[field];
  for(const field of ['fullTime','member'])if(['Yes','No'].includes(answers[field]))safe[field]=answers[field];
  for(let i=0;i<3;i++){const value=answers['graduates'+i];if(Number.isInteger(value)&&value>=0&&value<=1000000)safe['graduates'+i]=String(value);}
  return safe;
 }catch{checklistPrefillError='This personalized link is invalid or belongs to another scholarship year. Please request a current link or enter your details below.';return null;}
}
const checklistPrefill=readChecklistPrefill();
const prefillIdentity=checklistPrefill?checklistPrefill.university.trim().replace(/\s+/g,' ').toLowerCase():'blank';
const draftKey='phisim-university-qualification-draft-v2-'+qualificationDueYear+'-'+encodeURIComponent(prefillIdentity);
function applyChecklistPrefill(){
 if(checklistPrefill){
  for(const [field,value]of Object.entries(checklistPrefill)){const control=form.elements.namedItem(field);if(control&&field!=='year')control.value=value;}
  document.querySelector('#draft-status').textContent='Details pre-filled for '+checklistPrefill.university+'. Please review and update them before submitting.';
 }else if(checklistPrefillError)document.querySelector('#draft-status').textContent=checklistPrefillError;
}
// End PHISIM personalized form
'''


def install_prefill(js):
    if '// PHISIM personalized form v1' in js:
        js = re.sub(r'// PHISIM personalized form v1.*?// End PHISIM personalized form\s*', lambda _: PREFILL_JS, js, flags=re.S)
    else:
        js = replace_once(js, r"const draftKey='phisim-university-qualification-draft-v1';", PREFILL_JS, 'personalized form loader')
    if '\napplyChecklistPrefill();\nform.elements.year.value=qualificationDueYear;' not in js:
        js = replace_once(js, r'\nform\.elements\.year\.value=qualificationDueYear;',
                          '\napplyChecklistPrefill();\nform.elements.year.value=qualificationDueYear;', 'initial form prefill')
    return js

SUBMISSION_JS = r'''// PHISIM checklist submission v1
let submissionConfig=null, submissionKey=null, submittedSignature=null;
const submitButton=document.querySelector('#submit-to-checklist');
const submitStatus=document.querySelector('#submission-status');
form.addEventListener('input',()=>{if(submittedSignature&&JSON.stringify(data())!==submittedSignature){submitButton.disabled=false;submitButton.textContent='SUBMIT for Review to PHISIM';submitStatus.textContent='Your changes have not been submitted.';}});
submitButton.addEventListener('click',async()=>{
 trimValues();if(!form.reportValidity()){dialog.close();return;}
 const response=data(), signature=JSON.stringify(response);
 if(signature===submittedSignature){submitStatus.textContent='Your information is already saved and submitted for PHISIM review.';return;}
 submitButton.disabled=true;submitButton.textContent='Saving…';submitButton.setAttribute('aria-busy','true');
 submitStatus.textContent='Your information is being saved and submitted for review. Please wait…';
 try{
  if(location.protocol==='file:')throw new Error('Open this form through the local submission server. Run 02B_update_phisim_qualification_form.py --serve and use the address it displays.');
  if(!submissionConfig){
   const configReply=await fetch('/api/config',{cache:'no-store'});
   if(!configReply.ok)throw new Error('The submission server is unavailable. Start 02B_update_phisim_qualification_form.py --serve on the workbook computer.');
   submissionConfig=await configReply.json();
  }
  if(Number(response.year)!==submissionConfig.year)throw new Error('The scholarship year changed. Reload the form before submitting.');
  if(!submissionKey||submissionKey.signature!==signature)submissionKey={signature,id:crypto.randomUUID()};
  const result=await fetch('/api/qualification',{method:'POST',headers:{'Content-Type':'application/json','X-PHISIM-Submission':submissionConfig.token},body:JSON.stringify({...response,submissionId:submissionKey.id})});
  let saved;try{saved=await result.json();}catch{throw new Error('The server did not confirm a workbook save. Please try again.');}
  if(!result.ok)throw new Error(saved.error||'The response was not saved.');
  submittedSignature=signature;
  submitStatus.textContent=`Your information for ${saved.university} has been saved and submitted for PHISIM review.`;
 }catch(error){submitStatus.textContent=error.message||'The response was not saved. Please try again.';}
 finally{submitButton.removeAttribute('aria-busy');submitButton.disabled=submittedSignature===signature;submitButton.textContent=submittedSignature===signature?'Submitted for review':'SUBMIT for Review to PHISIM';}
});
// End PHISIM checklist submission
'''


def install_submission(html, js):
    """Make the annual updater preserve and install the real submit action."""
    html = html.replace('Your responses are not sent online. Review, then download or print your completed application to share with PHISIM.',
                        'Review your responses, then select Submit application to save them to the university evaluation checklist.')
    html = html.replace('This review is not an approval or an online submission.',
                        'Review your answers, then select Submit application. Submission does not constitute approval.')
    if 'id="submit-to-checklist"' not in html:
        html = replace_once(html, r'<div class="dialog-actions">',
                            '<p id="submission-status" role="status" aria-live="polite" class="hint"></p><div class="dialog-actions"><button id="submit-to-checklist" type="button" class="primary">Submit application</button>', 'submission controls')
    html = html.replace('Submit application', 'Submit for review')
    html = html.replace('Review your responses, then select Submit for review to save them to the university evaluation checklist.',
                        'Review your submission and click on the SUBMIT for Review to PHISIM')
    html = html.replace('Submit for review', 'SUBMIT for Review to PHISIM')
    js = js.replace('This document has not been submitted online. Send it to your PHISIM contact for review. Questions: president@phisim.org.',
                    'Use Submit application in the web form to save your responses to the PHISIM checklist. A downloaded copy alone does not submit the application. Questions: president@phisim.org.')
    js = js.replace('Use Submit application in the web form', 'Use Submit for review in the web form')
    js = re.sub(r'// PHISIM checklist submission v1.*?// End PHISIM checklist submission\s*', '', js, flags=re.S)
    return html, js.rstrip() + '\n' + SUBMISSION_JS


def organization_key(value):
    return ' '.join(str(value or '').split()).casefold()


def validate_submission(payload, year):
    if not isinstance(payload, dict):
        raise ValueError('Submission must be a JSON object.')
    limits = {'university': 200, 'school': 200, 'programs': 3000,
              'accreditation': 2000, 'faculty': 200, 'email': 254}
    result = {}
    for field, maximum in limits.items():
        value = payload.get(field, '')
        if not isinstance(value, str) or len(value) > maximum:
            raise ValueError(f'Invalid {field}.')
        result[field] = value.strip()
        if field != 'email' and not result[field]:
            raise ValueError(f'{field} is required.')
    if result['email'] and not re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+', result['email']):
        raise ValueError('Invalid faculty email.')
    parsed = urlsplit(result['accreditation'])
    if parsed.scheme not in ('https', 'http') or not parsed.netloc:
        raise ValueError('Accreditation webpage must be an HTTP or HTTPS URL.')
    for field in ('fullTime', 'member'):
        if payload.get(field) not in ('Yes', 'No'):
            raise ValueError(f'{field} must be Yes or No.')
        result[field] = payload[field]
    if str(payload.get('year')) != str(year):
        raise ValueError('Submission year does not match the active scholarship year. Reload the form.')
    result['year'] = year
    for index in range(3):
        field = f'graduates{index}'
        value = str(payload.get(field, ''))
        if not re.fullmatch(r'\d{1,7}', value) or not 0 <= int(value) <= 1000000:
            raise ValueError(f'{field} must be a whole number from 0 to 1000000.')
        result[field] = int(value)
    submission_id = payload.get('submissionId', '')
    if not isinstance(submission_id, str) or not re.fullmatch(r'[A-Za-z0-9-]{16,80}', submission_id):
        raise ValueError('Invalid submission identifier. Reload the form and try again.')
    result['submissionId'] = submission_id
    return result


def save_submission(workbook_path, payload, year):
    with WORKBOOK_WRITE_LOCK:
        return _save_submission(workbook_path, payload, year)


def _save_submission(workbook_path, payload, year):
    """Update exactly one existing university, keeping review decisions intact."""
    import openpyxl
    from copy import copy
    data = validate_submission(payload, year)
    workbook_path = Path(workbook_path)
    book = openpyxl.load_workbook(workbook_path)
    temporary = None
    try:
        candidates = [s for s in book if organization_key(s.title) == 'current evaluation checklist']
        if len(candidates) != 1:
            raise ValueError('The workbook needs exactly one Current evaluation checklist tab.')
        sheet = candidates[0]
        def header_key(value):
            return re.sub(r'[^a-z0-9]', '', str(value or '').casefold())
        header_rows = [row for row in sheet.iter_rows(min_row=1, max_row=min(30, sheet.max_row))
                       if any(header_key(c.value) == 'memberorganization' for c in row)]
        if len(header_rows) != 1:
            raise ValueError('Cannot identify a unique Member-Organization header row.')
        header_row = header_rows[0][0].row
        headers = {}
        for cell in header_rows[0]:
            if cell.value is not None:
                key = header_key(cell.value)
                if key in headers:
                    raise ValueError(f'Duplicate checklist header: {cell.value}')
                headers[key] = cell.column
        school_col = headers['memberorganization']
        matches = [r for r in range(header_row + 1, sheet.max_row + 1)
                   if organization_key(sheet.cell(r, school_col).value) == organization_key(data['university'])]
        if len(matches) != 1:
            raise ValueError('University name must match exactly one Member-Organization row (ignoring capitalization and extra spaces). '
                             f'Found {len(matches)} matches for {data["university"]!r}.')
        row = matches[0]
        title_years = re.findall(r'\b20\d{2}\b', str(sheet['A1'].value or ''))
        if not title_years or int(title_years[0]) != year:
            raise ValueError('Checklist year does not match the active form year. Update the checklist first.')
        id_col = headers.get('formsubmissionid')
        if id_col and sheet.cell(row, id_col).value == data['submissionId']:
            return {'university': sheet.cell(row, school_col).value, 'row': row, 'alreadySaved': True}
        updates = {
            'Submitted school or college': data['school'], 'Major': data['programs'],
            'Accreditation webpage': data['accreditation'],
            'Faculty representative': data['faculty'], 'Submitted faculty email': data['email'],
            'Full-time faculty member': data['fullTime'],
            'Current SIM Philadelphia academic member': data['member'],
            'Academic-year sequence ends in': year,
            **{f'Graduates {year-i-1}–{year-i}': data[f'graduates{i}'] for i in range(3)},
            'Received university materials': datetime.now(), 'Form submission ID': data['submissionId'],
            'Submission review status': 'Submitted for review',
        }
        for label, value in updates.items():
            key = header_key(label)
            col = headers.get(key)
            if col is None:
                col = sheet.max_column + 1
                headers[key] = col
                target = sheet.cell(header_row, col, label)
                template = sheet.cell(header_row, school_col)
                target._style = copy(template._style)
                target.alignment = copy(template.alignment)
                sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 28
            cell = sheet.cell(row, col)
            cell._style = copy(sheet.cell(row, school_col)._style)
            cell.value = value
            # Submitted text is stored as literal text, including leading '='.
            if isinstance(value, str):
                cell.data_type = 's'
            if isinstance(value, datetime):
                cell.number_format = 'mm/dd/yyyy hh:mm'
        if sheet.auto_filter.ref:
            old = openpyxl.utils.range_boundaries(sheet.auto_filter.ref)
            sheet.auto_filter.ref = f'A{header_row}:{openpyxl.utils.get_column_letter(sheet.max_column)}{old[3]}'
        fd, name = tempfile.mkstemp(suffix='.xlsx', dir=workbook_path.parent)
        os.close(fd)
        temporary = Path(name)
        book.save(temporary)
        verify = openpyxl.load_workbook(temporary)
        try:
            if verify[sheet.title].cell(row, headers['formsubmissionid']).value != data['submissionId']:
                raise ValueError('Workbook save verification failed.')
        finally:
            verify.close()
        backup = workbook_path.with_name(workbook_path.stem + '.submission-backup-' + datetime.now().strftime('%Y%m%d-%H%M%S-%f') + '.xlsx')
        shutil.copy2(workbook_path, backup)
        os.replace(temporary, workbook_path)
        return {'university': sheet.cell(row, school_col).value, 'row': row, 'alreadySaved': False}
    finally:
        book.close()
        if temporary:
            temporary.unlink(missing_ok=True)


def read_sync_configuration(path):
    config = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    parsed = urlsplit(config['url'])
    if parsed.scheme != 'https' or parsed.hostname != 'phisim-university-qualification.righttune7.chatgpt.site':
        raise ValueError('Sync URL must be the PHISIM qualification site over HTTPS.')
    class Blob(ctypes.Structure):
        _fields_ = [('size', ctypes.c_ulong), ('data', ctypes.POINTER(ctypes.c_ubyte))]
    raw = bytes.fromhex(config['protectedToken'])
    buffer = (ctypes.c_ubyte * len(raw)).from_buffer_copy(raw)
    incoming, outgoing = Blob(len(raw), buffer), Blob()
    if not ctypes.windll.crypt32.CryptUnprotectData(ctypes.byref(incoming), None, None, None, None, 0, ctypes.byref(outgoing)):
        raise OSError('The sync credential can only be opened by the Windows user who created it.')
    try:
        token = ctypes.string_at(outgoing.data, outgoing.size).decode('utf-16-le')
    finally:
        ctypes.windll.kernel32.LocalFree(outgoing.data)
    return config['url'].rstrip('/'), token


def sync_public_submissions(args, year):
    base, token = read_sync_configuration(args.sync_config)
    headers = {'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json',
               'User-Agent': 'PHISIM-Workbook-Sync/1.0'}
    with urlopen(Request(base + '/api/review-sync', headers=headers), timeout=30) as reply:
        items = json.load(reply)['submissions']
    count = 0
    for item in items:
        # Acknowledgment is sent only after an atomic, verified workbook save.
        save_submission(args.workbook, item['payload'], year)
        acknowledgment = json.dumps({'id': item['id'], 'workbookSaved': True}).encode('utf-8')
        with urlopen(Request(base + '/api/review-sync', data=acknowledgment, headers=headers), timeout=30) as reply:
            if not json.load(reply).get('acknowledged'):
                raise ValueError('The workbook was saved but acknowledgment was not confirmed. A retry is safe.')
        count += 1
        print(f'Online submission saved for review: {item["payload"]["university"]}', flush=True)
    return count


def submission_server(args, year):
    """Serve only the form assets; never expose the workbook or directory listing."""
    token = secrets.token_urlsafe(32)
    save_lock = threading.Lock()
    allowed = {'index.html', 'app.js', 'style.css'}
    allowed.update(p.name for p in args.form_dir.glob('phisim-logo.*') if p.suffix.lower() in ('.jpg', '.jpeg', '.png'))
    class Handler(BaseHTTPRequestHandler):
        def reply(self, code, value):
            content = json.dumps(value).encode('utf-8')
            self.send_response(code)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Content-Length', str(len(content)))
            self.send_header('Cache-Control', 'no-store')
            self.end_headers()
            self.wfile.write(content)

        def allowed_host(self):
            return self.headers.get('Host') in (f'127.0.0.1:{self.server.server_port}', f'localhost:{self.server.server_port}')

        def do_GET(self):
            if not self.allowed_host():
                self.reply(403, {'error': 'Use the local server address.'}); return
            path = urlsplit(self.path).path
            if path == '/api/config':
                self.reply(200, {'year': year, 'token': token}); return
            name = 'index.html' if path == '/' else path.lstrip('/')
            if name not in allowed:
                self.reply(404, {'error': 'Not found.'}); return
            content = (args.form_dir / name).read_bytes()
            self.send_response(200)
            self.send_header('Content-Type', mimetypes.guess_type(name)[0] or 'application/octet-stream')
            self.send_header('Content-Length', str(len(content)))
            self.send_header('Cache-Control', 'no-store')
            self.end_headers()
            self.wfile.write(content)

        def do_POST(self):
            if urlsplit(self.path).path != '/api/qualification':
                self.reply(404, {'error': 'Not found.'}); return
            origin = self.headers.get('Origin')
            expected = f'http://{self.headers.get("Host", "")}'
            if not self.allowed_host() or origin != expected or not secrets.compare_digest(self.headers.get('X-PHISIM-Submission', ''), token):
                self.reply(403, {'error': 'Open the form through this local server before submitting.'}); return
            try:
                length = int(self.headers.get('Content-Length', '0'))
                if not 0 < length <= 65536 or self.headers.get('Content-Type', '').split(';')[0] != 'application/json':
                    raise ValueError('Expected a JSON form submission of at most 64 KB.')
                payload = json.loads(self.rfile.read(length))
                with save_lock:
                    result = save_submission(args.workbook, payload, year)
                self.reply(200, result)
            except PermissionError:
                self.reply(409, {'error': 'Workbook could not be saved. Close it in Excel, then submit again.'})
            except (ValueError, KeyError, TypeError, UnicodeError) as error:
                self.reply(400, {'error': str(error)})
            except OSError:
                self.reply(503, {'error': 'Workbook is unavailable. Check its path and folder permissions, then try again.'})
    return ThreadingHTTPServer(('127.0.0.1', args.port), Handler)


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    match = DATE_PATTERN.search(str(value or ''))
    if match:
        return datetime.strptime(' '.join(match.groups()), '%B %d %Y').date()
    try:
        return date.fromisoformat(str(value).strip())
    except ValueError:
        raise ValueError(f'Cannot read qualification date: {value!r}') from None


def docx_deadline(path):
    from docx import Document
    document = Document(path)
    for table in document.tables:
        for index, row in enumerate(table.rows):
            if 'university qualification' in row.cells[0].text.strip().casefold():
                for following in table.rows[index + 1:index + 3]:
                    for cell in following.cells:
                        if DATE_PATTERN.search(cell.text):
                            return parse_date(cell.text)
    raise ValueError('No dated University qualification row found.')


def xlsx_deadline(path):
    import openpyxl
    from openpyxl.utils.datetime import from_excel
    formulas = openpyxl.load_workbook(path, data_only=False)
    cached = openpyxl.load_workbook(path, data_only=True)
    try:
        for sheet in formulas:
            for row in sheet:
                for label in row:
                    if str(label.value or '').strip().casefold() != 'university qualification':
                        continue
                    cell = sheet.cell(label.row, label.column + 1)
                    value = cached[sheet.title][cell.coordinate].value
                    if value is not None:
                        if isinstance(value, (int, float)):
                            value = from_excel(value, formulas.epoch)
                        return parse_date(value)
                    expression = str(cell.value or '')
                    # Support the actual schedule calculator without evaluating arbitrary code.
                    match = re.fullmatch(r'=\s*(\$?[A-Z]+\$?\d+)\s*-\s*(\$?[A-Z]+\$?\d+|\d+)\s*', expression, re.I)
                    if match:
                        anchor = sheet[match[1].replace('$', '')].value
                        offset = int(match[2]) if match[2].isdigit() else sheet[match[2].replace('$', '')].value
                        if isinstance(anchor, (int, float)):
                            anchor = from_excel(anchor, formulas.epoch)
                        if not isinstance(offset, (int, float)) or offset < 0:
                            raise ValueError('Schedule lead time must be a nonnegative number.')
                        return parse_date(anchor) - timedelta(days=offset)
                    if not expression.startswith('='):
                        return parse_date(cell.value)
                    raise ValueError(f'Unsupported formula {expression!r}. Recalculate and save the schedule in Excel, then retry.')
        raise ValueError('No University qualification row found.')
    finally:
        formulas.close()
        cached.close()


def select_schedule(args):
    if args.schedule:
        candidates = [args.schedule.resolve()]
        match = VERSION.match(candidates[0].name)
        selected_year = args.year or (int(match[1]) if match else None)
    else:
        versions = [(int(m[1]), p) for p in args.schedule_dir.iterdir()
                    if p.is_file() and (m := VERSION.match(p.name))]
        if args.year:
            versions = [(year, p) for year, p in versions if year == args.year]
        if not versions:
            raise ValueError('No matching year-specific schedule found. Use --schedule to select a file.')
        selected_year = max(year for year, _ in versions)
        candidates = sorted((p for year, p in versions if year == selected_year),
                            key=lambda p: (p.suffix.lower() != '.docx', p.name))
    failures = []
    for path in candidates:
        try:
            if path.suffix.lower() == '.docx':
                due = docx_deadline(path)
            elif path.suffix.lower() == '.xlsx':
                due = xlsx_deadline(path)
            else:
                raise ValueError('Schedule must be .docx or .xlsx.')
            if selected_year and due.year != selected_year:
                raise ValueError(f'Deadline year {due.year} disagrees with selected schedule year {selected_year}.')
            for failure in failures:
                print('Alternate source used: ' + failure)
            return path, due
        except (OSError, ValueError, KeyError) as error:
            failures.append(f'{path.name}: {error}')
    raise ValueError('Unable to read the selected schedule year; older years were not substituted.\n' + '\n'.join(failures))


def replace_once(text, pattern, replacement, name, flags=0):
    result, count = re.subn(pattern, lambda match: replacement(match) if callable(replacement) else replacement,
                            text, flags=flags)
    if count != 1:
        raise ValueError(f'Expected one {name}; found {count}. Form layout may have changed. No files saved.')
    return result


def choose_logo(args, source):
    if args.logo:
        data = args.logo.read_bytes()
        suffix = args.logo.suffix.lower()
    else:
        existing = [p for p in args.form_dir.glob('phisim-logo.*') if p.suffix.lower() in ('.jpg', '.jpeg', '.png')]
        if len(existing) == 1:
            return existing[0].name, existing[0].read_bytes()
        # Extract only an unambiguous image; never guess between several graphics.
        documents = [source] if source.suffix.lower() == '.docx' else []
        documents += sorted(args.schedule_dir.glob('*PHISIM Scholarship Process Dates.docx'))
        for document in dict.fromkeys(documents):
            try:
                with ZipFile(document) as archive:
                    images = [n for n in archive.namelist() if n.startswith('word/media/')
                              and Path(n).suffix.lower() in ('.jpg', '.jpeg', '.png')]
                    if len(images) == 1:
                        suffix = Path(images[0]).suffix.lower()
                        data = archive.read(images[0])
                        break
            except OSError:
                continue
        else:
            raise ValueError('PHISIM logo not found. Supply --logo with the original JPG or PNG.')
    if suffix not in ('.jpg', '.jpeg', '.png'):
        raise ValueError('Logo must be an original JPG or PNG file.')
    return 'phisim-logo' + suffix, data


def transform(args, source, due):
    html = (args.form_dir / 'index.html').read_text(encoding='utf-8')
    js = (args.form_dir / 'app.js').read_text(encoding='utf-8')
    statement = 'Please complete your response by ' + due.strftime('%A, %B %d, %Y') + '.'
    html = re.sub(r'<p\b[^>]*>\s*ACADEMIC PARTNERS\s*</p>', '', html, flags=re.I)
    html = replace_once(html, r'<title>[^<]*</title>', f'<title>{TITLE}</title>', 'browser title')
    html = replace_once(html, r'(<span class="header-note">)[^<]*(</span>)',
                        lambda m: m[1] + TITLE + m[2], 'header title')
    html = re.sub(r'<h2>(?:Your application|(?:PHISIM )?University Qualification(?: Form)?)</h2>',
                  f'<h2>{TITLE}</h2>', html, flags=re.I)
    html = replace_once(html, r'(<h2 id="review-title">)[^<]*(</h2>)',
                        lambda m: m[1] + TITLE + m[2], 'review title')
    html = replace_once(html, r'<p class="qualification-deadline"><strong>.*?</strong></p>',
                        f'<p class="qualification-deadline"><strong>{statement}</strong></p>', 'deadline', re.S)
    js = replace_once(js, r'const qualificationDeadlineStatement="[^"\n]*";',
                      'const qualificationDeadlineStatement=' + json.dumps(statement) + ';', 'JavaScript deadline')
    if 'const qualificationDueYear=' not in js or 'const y=qualificationDueYear;' not in js:
        raise ValueError('This script requires the deadline-based academic-year version of app.js.')
    html = replace_once(html, r'<input\b[^>]*\bid="year"[^>]*>',
        f'<input id="year" name="year" type="number" min="2000" max="2100" value="{due.year}" readonly aria-readonly="true">', 'academic-year input')
    for index in range(3):
        html = replace_once(html, rf'(<span class="year-label" id="year{index}">)[^<]*(</span>)',
            lambda m, i=index: m[1] + f'{due.year-i-1}–{due.year-i}' + m[2], f'academic-year label {index}')
    js = re.sub(r'<title>PHISIM University Qualification(?: Form)? —', f'<title>{TITLE} —', js)
    js = re.sub(r'<h1>PHISIM University Qualification(?: Form)?</h1>', f'<h1>{TITLE}</h1>', js)
    logo_name, logo_data = choose_logo(args, source)
    mime = mimetypes.guess_type(logo_name)[0]
    uri = 'data:' + mime + ';base64,' + base64.b64encode(logo_data).decode('ascii')
    js = replace_once(js, r"const phisimLogoDataUrl='[^']*';", "const phisimLogoDataUrl='" + uri + "';", 'download logo')
    html = re.sub(r'src="phisim-logo\.(?:jpg|jpeg|png)"', 'src="' + logo_name + '"', html)
    if 'ACADEMIC PARTNERS' in html.upper() or statement not in js or TITLE not in html:
        raise ValueError('Updated form failed verification. No files saved.')
    html, js = install_submission(html, install_prefill(js))
    css = (args.form_dir / 'style.css').read_text(encoding='utf-8')
    if 'PHISIM visible review actions v1' not in css:
        css = css.rstrip() + '\n' + REVIEW_DIALOG_CSS
    return {'index.html': html.encode('utf-8'), 'app.js': js.encode('utf-8'),
            'style.css': css.encode('utf-8'), logo_name: logo_data}, statement


def run(args):
    args.form_dir = args.form_dir.resolve()
    args.schedule_dir = args.schedule_dir.resolve()
    if not (args.form_dir / 'style.css').is_file():
        raise FileNotFoundError('Form folder must contain index.html, app.js, style.css, and the logo.')
    source, due = select_schedule(args)
    files, statement = transform(args, source, due)
    changed = {name: data for name, data in files.items()
               if not (args.form_dir / name).exists() or (args.form_dir / name).read_bytes() != data}
    print(f'Schedule: {source}')
    print(statement)
    print('Academic years: ' + ', '.join(f'{due.year-i-1}–{due.year-i}' for i in range(3)))
    print('Changed files: ' + (', '.join(changed) or 'none; form is already up to date'))
    if args.dry_run:
        print('Preview only. No files changed.')
        return
    if not changed:
        return
    backups = args.form_dir.parent / (args.form_dir.name + ' Backups')
    backup = backups / datetime.now().strftime('%Y%m%d-%H%M%S-%f')
    shutil.copytree(args.form_dir, backup)
    with tempfile.TemporaryDirectory(prefix='phisim-form-', dir=args.form_dir.parent) as temp:
        stage = Path(temp)
        for name, data in changed.items():
            (stage / name).write_bytes(data)
        try:
            for name in changed:
                os.replace(stage / name, args.form_dir / name)
        except OSError:
            for name in changed:
                previous = backup / name
                if previous.exists():
                    shutil.copy2(previous, args.form_dir / name)
                elif (args.form_dir / name).exists():
                    (args.form_dir / name).unlink()
            raise
    print(f'Backup: {backup}')
    print(f'Saved local form: {args.form_dir / "index.html"}')
    print('Hosted website unchanged. Publish the updated form separately to update its public URL.')


def main():
    folder = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--form-dir', type=Path, default=folder / 'University Qualification Web Form')
    parser.add_argument('--schedule-dir', type=Path, default=folder.parent / SCHEDULE_FOLDER)
    parser.add_argument('--schedule', type=Path, help='Use this exact DOCX or XLSX schedule instead of auto-selection.')
    parser.add_argument('--year', type=int, help='Choose a specific schedule year; default chooses the latest.')
    parser.add_argument('--logo', type=Path, help='Optional replacement original PHISIM JPG or PNG.')
    parser.add_argument('--dry-run', action='store_true', help='Preview updates without saving or making backups.')
    parser.add_argument('--serve', action='store_true', help='Serve the local form and save submitted responses to the checklist.')
    parser.add_argument('--port', type=int, default=8768, help='Local submission server port (default: 8768).')
    parser.add_argument('--workbook', type=Path, default=folder / 'University Qualification Tracking requirements.xlsx',
                        help='Workbook to update when the local form is submitted.')
    parser.add_argument('--sync-public', action='store_true', help='Copy queued public form submissions to the matching workbook rows, then exit.')
    parser.add_argument('--sync-config', type=Path, default=folder / 'PHISIM Submission Sync.json', help='Windows-protected public submission sync configuration.')
    args = parser.parse_args()
    if args.serve and args.dry_run:
        parser.error('--serve cannot be combined with --dry-run')
    if not 1 <= args.port <= 65535:
        parser.error('--port must be from 1 to 65535')
    try:
        run(args)
        if args.sync_public:
            _, due = select_schedule(args)
            count = sync_public_submissions(args, due.year)
            print(f'Online submissions synchronized: {count}')
        if args.serve:
            _, due = select_schedule(args)
            server = submission_server(args, due.year)
            print(f'Open the form: http://127.0.0.1:{server.server_port}/', flush=True)
            print(f'Submissions update: {args.workbook}', flush=True)
            print('Keep this window running. Press Ctrl+C to stop.', flush=True)
            sync_stop = threading.Event()
            if args.sync_config.exists():
                def sync_loop():
                    while not sync_stop.is_set():
                        try:
                            sync_public_submissions(args, due.year)
                        except Exception as error:
                            print(f'Online submission sync will retry: {error}', flush=True)
                        sync_stop.wait(30)
                threading.Thread(target=sync_loop, daemon=True).start()
                print('Public submissions sync to the workbook every 30 seconds while this server runs.', flush=True)
            try:
                server.serve_forever()
            except KeyboardInterrupt:
                print('\nSubmission server stopped.')
            finally:
                sync_stop.set()
                server.server_close()
    except ImportError as error:
        parser.exit(1, f'{error}\nInstall dependencies with: py -m pip install openpyxl python-docx\n')
    except (OSError, ValueError, KeyError) as error:
        parser.exit(1, f'{error}\nNo older schedule year was substituted. Close open source files if they are locked.\n')


if __name__ == '__main__':
    main()
