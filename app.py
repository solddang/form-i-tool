from flask import Flask, request, send_file, render_template_string
import zipfile, io, os, subprocess, tempfile

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'FormI_HS2026-05-302P_Template.docx')

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
NEW_FIXURE_BLOCK = (
    '<w:r><w:rPr><w:sz w:val="23"/><w:szCs w:val="23"/><w:lang w:eastAsia="ko-KR"/></w:rPr>'
    '<w:t>CHECKING FIXURE: 9031.80</w:t></w:r>'
)

HTML = open(os.path.join(BASE_DIR, 'index.html'), encoding='utf-8').read()

def build_docx(desc):
    new_xml = desc.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    with zipfile.ZipFile(TEMPLATE_PATH, 'r') as z:
        files = {name: z.read(name) for name in z.namelist()}
    xml = files['word/document.xml'].decode('utf-8')
    xml = xml.replace(OLD_DESC, new_xml)
    xml = xml.replace(OLD_FIXURE_BLOCK, NEW_FIXURE_BLOCK)
    files['word/document.xml'] = xml.encode('utf-8')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    buf.seek(0)
    return buf

def docx_to_pdf(docx_buf):
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, 'input.docx')
        with open(docx_path, 'wb') as f:
            f.write(docx_buf.read())
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', tmpdir, docx_path],
            capture_output=True, timeout=60
        )
        pdf_path = os.path.join(tmpdir, 'input.pdf')
        if result.returncode != 0 or not os.path.exists(pdf_path):
            raise RuntimeError('PDF 변환 실패: ' + result.stderr.decode())
        with open(pdf_path, 'rb') as f:
            return io.BytesIO(f.read())

def has_libreoffice():
    try:
        subprocess.run(['libreoffice', '--version'], capture_output=True, timeout=5)
        return True
    except:
        return False

@app.route('/')
def index():
    return HTML

@app.route('/status')
def status():
    return {'libreoffice': has_libreoffice()}

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    no, desc, fmt = int(data['no']), data['desc'], data.get('fmt', 'docx')
    docx_buf = build_docx(desc)
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
            no, desc = int(item['no']), item['desc']
            docx_buf = build_docx(desc)
            if fmt == 'pdf':
                pdf_buf = docx_to_pdf(docx_buf)
                zout.writestr(f"FormI_HS2026-05-302P_NO {no:03d}.pdf", pdf_buf.read())
            else:
                zout.writestr(f"FormI_HS2026-05-302P_NO {no:03d}.docx", docx_buf.read())
    zip_buf.seek(0)
    return send_file(zip_buf, as_attachment=True, download_name=f'FormI_HS2026-05-302P_ALL.zip',
                     mimetype='application/zip')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"서버 시작: http://localhost:{port}")
    app.run(host='0.0.0.0', port=port)
