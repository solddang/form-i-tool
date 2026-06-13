from flask import Flask, request, send_file, Response
import zipfile, io, os, subprocess, tempfile, shutil

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'FormI_HS2026-05-302P_Template.docx')
SAMPLE_XLSX_PATH = os.path.join(BASE_DIR, 'sample_template.xlsx')

OLD_DESC = 'INSPECTION FIXTURE FOR PNL ASSY-CTR FLR COMPL OF PNL &amp; MBR ASSY CTR FLR COMPL (RHD/LHD) CONSISTS OF I/F FOR BC4i MODEL'
OLD_FIXURE_BLOCK = (
    '<w:r><w:rPr><w:sz w:val="23"/><w:szCs w:val="23"/><w:lang w:eastAsia="ko-KR"/></w:rPr>'
    '<w:t>INSPECTION</w:t></w:r>'
    '<w:r w:rsidR="00D32FF5" w:rsidRPr="007678EE"><w:rPr><w:sz w:val="23"/><w:szCs w:val="23"/><w:lang w:eastAsia="ko-KR"/></w:rPr>'
    '<w:t xml:space="preserve"> FIXURE</w:t></w:r>'
    '<w:r w:rsidR="00214A27" w:rsidRPr="007678EE"><w:rPr><w:sz w:val="23"/><w:szCs w:val="23"/><w:lang w:eastAsia="ko-KR"/></w:rPr>'
    '<w:t xml:space="preserve">: </w:t></w:r>'
    '<w:r w:rsidR="00D32FF5" w:rsidRPr="007678EE"><w:rPr><w:sz w:val="23"/><w:szCs w:val="23"/><w:lang w:eastAsia="ko-KR"/></w:rPr>'
    '<w:t>9031.80</w:t></w:r>'
)

CATEGORY_MAP = {
    'CHECKING FIXTURE': ('CHECKING FIXURE', '9031.80'),
    'PRESS TOOL':       ('PRESS TOOL',      '8207.30'),
    'JIG':              ('JIG',             '8466.30'),
}

def get_libreoffice():
    for cmd in ['libreoffice', 'soffice']:
        path = shutil.which(cmd)
        if path:
            return path
    for root, dirs, files in os.walk('/nix/store'):
        for f in files:
            if f in ('libreoffice', 'soffice'):
                full = os.path.join(root, f)
                if os.access(full, os.X_OK):
                    return full
        dirs[:] = [d for d in dirs if not d.startswith('.')]
    raise RuntimeError('LibreOffice를 찾을 수 없습니다.')

HTML = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Form I 자동 생성기 — HWASHIN</title>
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Malgun Gothic", sans-serif; background: #f4f5f7; color: #1a1a1a; min-height: 100vh; display: flex; align-items: flex-start; justify-content: center; padding: 2.5rem 1rem 3rem; }
.card { background: #fff; border-radius: 14px; border: 1px solid #e4e4e7; box-shadow: 0 2px 12px rgba(0,0,0,0.06); padding: 2rem; width: 100%; max-width: 720px; }
.header { margin-bottom: 1.75rem; }
.header-row { display: flex; align-items: center; gap: 10px; margin-bottom: 4px; }
.logo { width: 34px; height: 34px; border-radius: 9px; background: #1a56db; display: flex; align-items: center; justify-content: center; color: #fff; font-size: 17px; flex-shrink: 0; }
h1 { font-size: 18px; font-weight: 700; }
.sub { font-size: 12px; color: #9ca3af; padding-left: 44px; }
.info-bar { display: flex; flex-wrap: wrap; gap: 6px 18px; background: #f8faff; border: 1px solid #dbeafe; border-radius: 9px; padding: 10px 14px; margin-bottom: 1.5rem; font-size: 12px; color: #6b7280; }
.info-bar b { color: #1e40af; font-weight: 600; }

.cat-table { width: 100%; border-collapse: collapse; margin-bottom: 1.5rem; font-size: 12px; }
.cat-table th { background: #f1f5f9; color: #475569; font-weight: 600; padding: 7px 12px; border: 1px solid #e4e4e7; text-align: left; }
.cat-table td { padding: 6px 12px; border: 1px solid #e4e4e7; color: #374151; }
.cat-table tr:nth-child(even) td { background: #fafafa; }
.badge-cat { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
.badge-cf { background: #dbeafe; color: #1e40af; }
.badge-pt { background: #fef9c3; color: #854d0e; }
.badge-jig { background: #dcfce7; color: #166534; }

.sample-box { display: flex; align-items: center; justify-content: space-between; background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 9px; padding: 10px 14px; margin-bottom: 1.5rem; }
.sample-left { display: flex; align-items: center; gap: 10px; }
.sample-icon { font-size: 20px; }
.sample-text { font-size: 13px; color: #166534; font-weight: 500; }
.sample-hint { font-size: 11px; color: #4ade80; margin-top: 2px; color: #15803d; }
.sample-btn { display: inline-flex; align-items: center; gap: 6px; padding: 7px 14px; font-size: 12px; font-weight: 600; background: #16a34a; color: #fff; border: none; border-radius: 7px; cursor: pointer; text-decoration: none; }
.sample-btn:hover { background: #15803d; }

.fmt-row { display: flex; align-items: center; gap: 10px; margin-bottom: 1.25rem; }
.fmt-label { font-size: 12px; font-weight: 600; color: #6b7280; flex-shrink: 0; }
.fmt-toggle { display: flex; border: 1px solid #e4e4e7; border-radius: 8px; overflow: hidden; }
.fmt-btn { padding: 6px 16px; font-size: 12px; font-weight: 500; background: #fff; color: #6b7280; border: none; cursor: pointer; }
.fmt-btn:first-child { border-right: 1px solid #e4e4e7; }
.fmt-btn.active { background: #1a56db; color: #fff; }
.field-label { display: block; font-size: 12px; font-weight: 600; color: #374151; margin-bottom: 7px; }
textarea { width: 100%; font-size: 13px; font-family: "Consolas","Menlo",monospace; padding: 11px 13px; border: 1px solid #d1d5db; border-radius: 9px; background: #fff; color: #111; resize: vertical; min-height: 200px; line-height: 1.7; outline: none; }
textarea:focus { border-color: #1a56db; box-shadow: 0 0 0 3px rgba(26,86,219,0.1); }
.hint { font-size: 11px; color: #9ca3af; margin-top: 5px; line-height: 1.6; }
.btn { display: inline-flex; align-items: center; justify-content: center; gap: 7px; padding: 9px 16px; font-size: 13px; font-weight: 500; border-radius: 8px; border: 1px solid #d1d5db; background: #fff; color: #374151; cursor: pointer; }
.btn:hover:not(:disabled) { background: #f9fafb; }
.btn:disabled { opacity: 0.38; cursor: not-allowed; }
.btn-parse { width: 100%; margin-top: 10px; }
.btn-primary { background: #1a56db; color: #fff; border-color: #1a56db; font-weight: 600; }
.btn-primary:hover:not(:disabled) { background: #1648c0; }
.btn-secondary { background: #fff; color: #1a56db; border-color: #bfdbfe; }
.btn-secondary:hover:not(:disabled) { background: #eff6ff; }
.divider { border: none; border-top: 1px solid #f3f4f6; margin: 1.5rem 0; }
.preview-wrap { margin-top: 1.25rem; display: none; }
.preview-top { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; }
.preview-title { font-size: 12px; font-weight: 600; color: #374151; }
.badge { font-size: 11px; padding: 2px 10px; border-radius: 20px; font-weight: 600; }
.badge.ok { background: #dcfce7; color: #166534; }
.badge.warn { background: #fef9c3; color: #854d0e; }
.preview-list { border: 1px solid #e4e4e7; border-radius: 9px; overflow: hidden; max-height: 260px; overflow-y: auto; }
.p-row { display: grid; grid-template-columns: 36px 120px 1fr; gap: 8px; padding: 7px 13px; border-bottom: 1px solid #f3f4f6; font-size: 12px; align-items: start; }
.p-row:last-child { border-bottom: none; }
.p-row.err-row { background: #fff5f5; grid-template-columns: 1fr; }
.p-no { font-family: monospace; color: #9ca3af; }
.p-desc { color: #111; line-height: 1.5; }
.p-err { color: #dc2626; }
.dl-label { font-size: 12px; font-weight: 600; color: #374151; margin-bottom: 10px; }
.dl-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
.progress { height: 4px; background: #e4e4e7; border-radius: 4px; margin-top: 12px; display: none; overflow: hidden; }
.progress-fill { height: 100%; background: #1a56db; border-radius: 4px; width: 0%; transition: width 0.15s; }
.status { font-size: 12px; text-align: center; margin-top: 9px; min-height: 18px; color: #9ca3af; }
.status.ok { color: #16a34a; font-weight: 500; }
.status.err { color: #dc2626; }
.status.working { color: #1a56db; }
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <div class="header-row">
      <div class="logo">📋</div>
      <h1>Form I 자동 생성기</h1>
    </div>
    <div class="sub">HWASHIN CO. LTD. &nbsp;·&nbsp; CEPA 원산지확인서 Section III</div>
  </div>

  <div class="info-bar">
    <span>🔒 템플릿 <b>HS2026-05-302P</b></span>
    <span>🔒 기타 항목 고정</span>
  </div>

  <table class="cat-table">
    <thead><tr><th>Category</th><th>HS Code</th><th>Part B 상단 자동 적용</th></tr></thead>
    <tbody>
      <tr><td><span class="badge-cat badge-cf">CHECKING FIXTURE</span></td><td>9031.80</td><td>CHECKING FIXURE: 9031.80</td></tr>
      <tr><td><span class="badge-cat badge-pt">PRESS TOOL</span></td><td>8207.30</td><td>PRESS TOOL: 8207.30</td></tr>
      <tr><td><span class="badge-cat badge-jig">JIG</span></td><td>8466.30</td><td>JIG: 8466.30</td></tr>
    </tbody>
  </table>

  <div class="sample-box">
    <div class="sample-left">
      <span class="sample-icon">📥</span>
      <div>
        <div class="sample-text">엑셀 입력 양식 샘플</div>
        <div class="sample-hint">NO. / Category / Description 형식 — 다운로드 후 작성해서 붙여넣으세요</div>
      </div>
    </div>
    <a href="/sample" download="Form_I_입력양식_샘플.xlsx" class="sample-btn">⬇ 샘플 다운로드</a>
  </div>

  <div class="fmt-row">
    <span class="fmt-label">출력 형식</span>
    <div class="fmt-toggle">
      <button class="fmt-btn active" id="fmt-docx" onclick="setFmt('docx')">📄 Word (.docx)</button>
      <button class="fmt-btn" id="fmt-pdf" onclick="setFmt('pdf')">📕 PDF</button>
    </div>
  </div>

  <label class="field-label" for="inp">번호 + Category + 품목 설명 목록</label>
  <textarea id="inp" placeholder="샘플 엑셀에서 NO. / Category / Description 세 열을 복사해서 붙여넣으세요.

1	PRESS TOOL	PRESS TOOL FOR BRKT-DR SCUFF MTG NO.1 ...
2	CHECKING FIXTURE	CHECKING FIXTURE FOR PLATE-FR BUMPER ...
3	JIG	W15 FRT L/ARM REINF WELDING JIG"></textarea>
  <div class="hint">탭으로 구분된 세 열 (번호 / Category / 품목명) · 엑셀에서 세 열 선택 후 복사 붙여넣기</div>
  <button class="btn btn-parse" onclick="parse()">목록 확인</button>

  <div class="preview-wrap" id="preview-wrap">
    <div class="preview-top">
      <span class="preview-title">파싱 결과</span>
      <span class="badge ok" id="badge">0개</span>
    </div>
    <div class="preview-list" id="preview-list"></div>
  </div>

  <hr class="divider">
  <div class="dl-label">다운로드</div>
  <div class="dl-grid">
    <button class="btn btn-secondary" id="btn-one" onclick="genOne()" disabled>📄 개별 파일 (1개씩)</button>
    <button class="btn btn-primary" id="btn-zip" onclick="genZip()" disabled>📦 전체 ZIP</button>
  </div>
  <div class="progress" id="prog"><div class="progress-fill" id="prog-fill"></div></div>
  <div class="status" id="status"></div>
</div>

<script>
let items = [], fmt = 'docx';
const CAT_BADGE = {
  'CHECKING FIXTURE': '<span class="badge-cat badge-cf">CHECKING FIXTURE</span>',
  'PRESS TOOL':       '<span class="badge-cat badge-pt">PRESS TOOL</span>',
  'JIG':              '<span class="badge-cat badge-jig">JIG</span>',
};
function setFmt(f) {
  fmt = f;
  document.getElementById('fmt-docx').classList.toggle('active', f==='docx');
  document.getElementById('fmt-pdf').classList.toggle('active', f==='pdf');
}
function xesc(s) { return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function parse() {
  const lines = document.getElementById('inp').value.split('\\n');
  const ok = [], bad = [];
  lines.forEach((l, i) => {
    const t = l.trim(); if (!t) return;
    const parts = t.split('\\t');
    if (parts.length >= 3) {
      const no = parts[0].trim();
      const cat = parts[1].trim().toUpperCase();
      const desc = parts.slice(2).join('\\t').trim();
      if (/^\\d+$/.test(no) && ['CHECKING FIXTURE','PRESS TOOL','JIG'].includes(cat) && desc) {
        ok.push({ no: parseInt(no), cat, desc });
      } else { bad.push(i + 1); }
    } else { bad.push(i + 1); }
  });
  items = ok;
  const wrap=document.getElementById('preview-wrap'), list=document.getElementById('preview-list'), badge=document.getElementById('badge');
  if (!ok.length && !bad.length) { setStatus('내용을 입력해 주세요.', 'err'); return; }
  wrap.style.display = 'block';
  badge.textContent = ok.length + '개';
  badge.className = 'badge ' + (bad.length ? 'warn' : 'ok');
  list.innerHTML = ok.map(it =>
    `<div class="p-row"><span class="p-no">${it.no}</span><span>${CAT_BADGE[it.cat]||it.cat}</span><span class="p-desc">${xesc(it.desc)}</span></div>`
  ).join('') + (bad.length ? `<div class="p-row err-row"><span class="p-err">⚠ ${bad.length}줄 인식 실패 — Category 확인 필요</span></div>` : '');
  document.getElementById('btn-one').disabled = ok.length === 0;
  document.getElementById('btn-zip').disabled = ok.length === 0;
  setStatus(ok.length ? ok.length + '개 항목 준비됨' : '', ok.length ? 'ok' : '');
}
function fname(no) { return `FormI_HS2026-05-302P_NO ${String(no).padStart(3,'0')}.${fmt==='pdf'?'pdf':'docx'}`; }
function setStatus(msg, cls) { const e=document.getElementById('status'); e.textContent=msg; e.className='status '+(cls||''); }
function setProgress(pct) { const p=document.getElementById('prog'); p.style.display=(pct>0&&pct<100)?'block':'none'; document.getElementById('prog-fill').style.width=pct+'%'; }
function lock(v) { document.getElementById('btn-one').disabled=v; document.getElementById('btn-zip').disabled=v; }
async function genOne() {
  if (!items.length) return;
  lock(true); setProgress(5); setStatus('생성 중...', 'working');
  try {
    for (let i=0; i<items.length; i++) {
      const {no,cat,desc} = items[i];
      const r = await fetch('/generate', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({no,cat,desc,fmt})});
      if (!r.ok) throw new Error(await r.text());
      const blob = await r.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href=url; a.download=fname(no); a.click();
      URL.revokeObjectURL(url);
      setProgress(10 + Math.round((i+1)/items.length*85));
      await new Promise(r => setTimeout(r, 300));
    }
    setStatus('✓ ' + items.length + '개 다운로드 완료', 'ok');
  } catch(e) { setStatus('오류: ' + e.message, 'err'); }
  setProgress(0); lock(false);
}
async function genZip() {
  if (!items.length) return;
  lock(true); setProgress(10);
  setStatus(fmt==='pdf' ? 'PDF 변환 중... (시간이 걸릴 수 있어요)' : 'ZIP 생성 중...', 'working');
  try {
    const r = await fetch('/generate_zip', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({items,fmt})});
    if (!r.ok) throw new Error(await r.text());
    setProgress(90);
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href=url; a.download='FormI_HS2026-05-302P_ALL.zip'; a.click();
    URL.revokeObjectURL(url);
    setStatus('✓ ' + items.length + '개 ZIP 다운로드 완료', 'ok');
  } catch(e) { setStatus('오류: ' + e.message, 'err'); }
  setProgress(0); lock(false);
}
</script>
</body>
</html>"""

def make_fixure_block(label, hs_code):
    return (
        '<w:r><w:rPr><w:sz w:val="23"/><w:szCs w:val="23"/><w:lang w:eastAsia="ko-KR"/></w:rPr>'
        f'<w:t>{label}: {hs_code}</w:t></w:r>'
    )

def build_docx(desc, cat):
    label, hs_code = CATEGORY_MAP.get(cat, ('CHECKING FIXURE', '9031.80'))
    new_desc_xml = desc.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    with zipfile.ZipFile(TEMPLATE_PATH, 'r') as z:
        files = {name: z.read(name) for name in z.namelist()}
    xml = files['word/document.xml'].decode('utf-8')
    xml = xml.replace(OLD_DESC, new_desc_xml)
    xml = xml.replace(OLD_FIXURE_BLOCK, make_fixure_block(label, hs_code))
    files['word/document.xml'] = xml.encode('utf-8')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    buf.seek(0)
    return buf

def docx_to_pdf(docx_buf):
    lo = get_libreoffice()
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, 'input.docx')
        with open(docx_path, 'wb') as f:
            f.write(docx_buf.read())
        env = os.environ.copy()
        env['HOME'] = tmpdir
        result = subprocess.run(
            [lo, '--headless', '--norestore', '--convert-to', 'pdf', '--outdir', tmpdir, docx_path],
            capture_output=True, timeout=120, env=env
        )
        pdf_path = os.path.join(tmpdir, 'input.pdf')
        if result.returncode != 0 or not os.path.exists(pdf_path):
            raise RuntimeError(f'PDF 변환 실패: {result.stderr.decode()[:200]}')
        with open(pdf_path, 'rb') as f:
            return io.BytesIO(f.read())

@app.route('/')
def index():
    return Response(HTML, mimetype='text/html; charset=utf-8')

@app.route('/sample')
def sample():
    return send_file(SAMPLE_XLSX_PATH, as_attachment=True,
                     download_name='Form_I_입력양식_샘플.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    no, cat, desc, fmt = int(data['no']), data['cat'], data['desc'], data.get('fmt', 'docx')
    docx_buf = build_docx(desc, cat)
    if fmt == 'pdf':
        pdf_buf = docx_to_pdf(docx_buf)
        fname = f"FormI_HS2026-05-302P_NO {no:03d}.pdf"
        return send_file(pdf_buf, as_attachment=True, download_name=fname, mimetype='application/pdf')
    fname = f"FormI_HS2026-05-302P_NO {no:03d}.docx"
    return send_file(docx_buf, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.route('/generate_zip', methods=['POST'])
def generate_zip():
    data = request.json
    items, fmt = data['items'], data.get('fmt', 'docx')
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in items:
            no, cat, desc = int(item['no']), item['cat'], item['desc']
            docx_buf = build_docx(desc, cat)
            if fmt == 'pdf':
                pdf_buf = docx_to_pdf(docx_buf)
                zout.writestr(f"FormI_HS2026-05-302P_NO {no:03d}.pdf", pdf_buf.read())
            else:
                zout.writestr(f"FormI_HS2026-05-302P_NO {no:03d}.docx", docx_buf.read())
    zip_buf.seek(0)
    return send_file(zip_buf, as_attachment=True, download_name='FormI_HS2026-05-302P_ALL.zip',
                     mimetype='application/zip')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"서버 시작: http://localhost:{port}")
    app.run(host='0.0.0.0', port=port)
