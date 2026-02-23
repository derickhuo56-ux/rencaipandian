import cgi
import io
import json
import os
import posixpath
import secrets
import urllib.parse
import zipfile
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import xml.etree.ElementTree as ET

ROOT = Path(__file__).parent
PUBLIC = ROOT / 'public'
DATA = ROOT / 'data'
UPLOADS = ROOT / 'uploads'
DATA.mkdir(exist_ok=True)
UPLOADS.mkdir(exist_ok=True)

ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin123')
ADMIN_TOKEN = os.environ.get('ADMIN_TOKEN', secrets.token_hex(16))

CONFIG_PATH = DATA / 'config.json'
REVIEWS_PATH = DATA / 'reviews.json'

DEFAULT_CONFIG = {
    'departments': ['技术部', '市场部', '人力资源部'],
    'employees': [
        {'id': 1, 'name': '小明', 'department': '技术部', 'level': 'P2'},
        {'id': 2, 'name': '小红', 'department': '技术部', 'level': 'P3'},
        {'id': 3, 'name': '小李', 'department': '市场部', 'level': 'P2'},
        {'id': 4, 'name': '小王', 'department': '人力资源部', 'level': 'P1'}
    ],
    'behaviorsByLevel': {
        'P1': ['按流程完成任务', '按时提交周报', '与同事保持良好协作'],
        'P2': ['独立推进项目模块', '主动识别并解决问题', '跨团队有效沟通', '持续复盘并优化'],
        'P3': ['制定团队执行策略', '指导与培养团队成员', '推进关键跨部门协作', '驱动业务结果达成']
    }
}


def ensure_file(path: Path, default):
    if not path.exists():
        path.write_text(json.dumps(default, ensure_ascii=False, indent=2), encoding='utf-8')


def read_json(path: Path):
    return json.loads(path.read_text(encoding='utf-8'))


def write_json(path: Path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


def col_name(i: int):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def xml_escape(text: str):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def sheet_xml(rows):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
    for r_idx, row in enumerate(rows, start=1):
        out.append(f'<row r="{r_idx}">')
        for c_idx, val in enumerate(row):
            cell = f'{col_name(c_idx)}{r_idx}'
            out.append(f'<c r="{cell}" t="inlineStr"><is><t>{xml_escape(str(val))}</t></is></c>')
        out.append('</row>')
    out.append('</sheetData></worksheet>')
    return ''.join(out)


def build_template_xlsx_bytes():
    employees = [
        ['部门', '员工姓名', '职级'],
        ['技术部', '小明', 'P2'],
        ['市场部', '小李', 'P2']
    ]
    behaviors = [
        ['职级', '工作行为'],
        ['P1', '按流程完成任务'],
        ['P1', '按时提交周报'],
        ['P2', '独立推进项目模块'],
        ['P2', '主动识别并解决问题']
    ]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
        z.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="员工配置" sheetId="1" r:id="rId1"/><sheet name="行为配置" sheetId="2" r:id="rId2"/></sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        z.writestr('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>')
        z.writestr('xl/worksheets/sheet1.xml', sheet_xml(employees))
        z.writestr('xl/worksheets/sheet2.xml', sheet_xml(behaviors))
    return buf.getvalue()


def parse_xlsx(file_path: Path):
    ns = {
        'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'rel': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'pkg': 'http://schemas.openxmlformats.org/package/2006/relationships'
    }
    with zipfile.ZipFile(file_path, 'r') as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('main:si', ns):
                t = ''.join(node.text or '' for node in si.iterfind('.//main:t', ns))
                shared.append(t)

        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rid_to_target = {rel.attrib['Id']: rel.attrib['Target'] for rel in rels.findall('pkg:Relationship', ns)}
        sheets = {}
        for sh in wb.findall('main:sheets/main:sheet', ns):
            name = sh.attrib['name']
            rid = sh.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            sheets[name] = 'xl/' + rid_to_target[rid]

        def read_sheet(sheet_name):
            if sheet_name not in sheets:
                return []
            root = ET.fromstring(z.read(sheets[sheet_name]))
            rows = []
            for row in root.findall('.//main:row', ns):
                values = []
                for c in row.findall('main:c', ns):
                    ctype = c.attrib.get('t')
                    v = ''
                    if ctype == 'inlineStr':
                        v = ''.join(node.text or '' for node in c.iterfind('.//main:t', ns))
                    else:
                        vnode = c.find('main:v', ns)
                        raw = vnode.text if vnode is not None else ''
                        if ctype == 's' and raw.isdigit():
                            idx = int(raw)
                            v = shared[idx] if idx < len(shared) else ''
                        else:
                            v = raw
                    values.append(v)
                rows.append(values)
            return rows

        return read_sheet('员工配置'), read_sheet('行为配置')


ensure_file(CONFIG_PATH, DEFAULT_CONFIG)
ensure_file(REVIEWS_PATH, [])


class Handler(BaseHTTPRequestHandler):
    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, file_path: Path, ctype='text/html; charset=utf-8', download_name=None):
        if not file_path.exists():
            self.send_error(404)
            return
        data = file_path.read_bytes()
        self.send_response(200)
        self.send_header('Content-Type', ctype)
        if download_name:
            self.send_header('Content-Disposition', f'attachment; filename="{download_name}"')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def parse_json_body(self):
        length = int(self.headers.get('Content-Length', '0'))
        raw = self.rfile.read(length) if length else b'{}'
        return json.loads(raw.decode('utf-8') or '{}')

    def auth_ok(self):
        auth = self.headers.get('Authorization', '')
        return auth == f'Bearer {ADMIN_TOKEN}'

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == '/':
            return self.send_file(PUBLIC / 'user.html')

        if path.startswith('/api/admin/') and not self.auth_ok():
            return self.send_json({'message': '未授权'}, 401)

        if path == '/api/public/departments':
            return self.send_json(read_json(CONFIG_PATH)['departments'])

        if path == '/api/public/employees':
            q = urllib.parse.parse_qs(parsed.query)
            department = q.get('department', [''])[0]
            config = read_json(CONFIG_PATH)
            employees = config['employees']
            if department:
                employees = [e for e in employees if e['department'] == department]
            return self.send_json([{'id': e['id'], 'name': e['name']} for e in employees])

        if path.startswith('/api/public/employee/') and path.endswith('/behaviors'):
            eid = path.split('/')[4]
            config = read_json(CONFIG_PATH)
            employee = next((e for e in config['employees'] if str(e['id']) == str(eid)), None)
            if not employee:
                return self.send_json({'message': '员工不存在'}, 404)
            behaviors = config['behaviorsByLevel'].get(employee['level'], [])
            return self.send_json({'employee': {'id': employee['id'], 'name': employee['name']}, 'behaviors': behaviors})

        if path == '/api/admin/reviews':
            records = read_json(REVIEWS_PATH)
            records.sort(key=lambda x: x['updatedAt'], reverse=True)
            return self.send_json(records)

        if path == '/api/admin/config':
            return self.send_json(read_json(CONFIG_PATH))

        if path == '/api/admin/template':
            data = build_template_xlsx_bytes()
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', 'attachment; filename="talent_template.xlsx"')
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        if path.startswith('/public/'):
            local = ROOT / path.lstrip('/')
            return self.send_file(local)

        # static from /public
        local = PUBLIC / posixpath.normpath(path.lstrip('/'))
        if local.exists() and local.is_file():
            ctype = 'text/html; charset=utf-8'
            if str(local).endswith('.js'):
                ctype = 'application/javascript; charset=utf-8'
            elif str(local).endswith('.css'):
                ctype = 'text/css; charset=utf-8'
            return self.send_file(local, ctype)

        self.send_error(404)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path == '/api/admin/login':
            body = self.parse_json_body()
            if body.get('password') == ADMIN_PASSWORD:
                return self.send_json({'token': ADMIN_TOKEN})
            return self.send_json({'message': '密码错误'}, 401)

        if path == '/api/public/reviews':
            body = self.parse_json_body()
            config = read_json(CONFIG_PATH)
            employee = next((e for e in config['employees'] if str(e['id']) == str(body.get('employeeId'))), None)
            if not employee:
                return self.send_json({'message': '员工不存在'}, 404)
            scores = body.get('behaviorScores') or []
            behaviors = config['behaviorsByLevel'].get(employee['level'], [])
            if len(scores) != len(behaviors):
                return self.send_json({'message': '行为评分项不完整'}, 400)
            try:
                scores = [float(s) for s in scores]
                perf = float(body.get('performanceScore'))
            except Exception:
                return self.send_json({'message': '分数格式错误'}, 400)
            if any(s < 1 or s > 5 for s in scores) or perf < 1 or perf > 5:
                return self.send_json({'message': '分数必须在1到5之间'}, 400)

            avg = sum(scores) / len(scores)
            records = read_json(REVIEWS_PATH)
            payload = {
                'employeeId': employee['id'],
                'employeeName': employee['name'],
                'department': employee['department'],
                'level': employee['level'],
                'behaviorAverage': avg,
                'performanceScore': perf,
                'updatedAt': datetime.utcnow().isoformat() + 'Z'
            }
            idx = next((i for i, r in enumerate(records) if r['employeeId'] == employee['id']), -1)
            if idx >= 0:
                records[idx] = payload
            else:
                records.append(payload)
            write_json(REVIEWS_PATH, records)
            return self.send_json({'message': '提交成功', 'payload': payload})

        if path.startswith('/api/admin/') and not self.auth_ok():
            return self.send_json({'message': '未授权'}, 401)

        if path == '/api/admin/config/upload':
            form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={
                'REQUEST_METHOD': 'POST',
                'CONTENT_TYPE': self.headers.get('Content-Type')
            })
            if 'file' not in form:
                return self.send_json({'message': '请上传文件'}, 400)
            file_item = form['file']
            temp = UPLOADS / f'upload-{secrets.token_hex(8)}.xlsx'
            with open(temp, 'wb') as f:
                f.write(file_item.file.read())
            try:
                emp_rows, behavior_rows = parse_xlsx(temp)
                if not emp_rows or not behavior_rows:
                    return self.send_json({'message': '模板错误'}, 400)
                emp_header = emp_rows[0]
                beh_header = behavior_rows[0]
                emp_idx = {name: i for i, name in enumerate(emp_header)}
                beh_idx = {name: i for i, name in enumerate(beh_header)}
                for key in ['部门', '员工姓名', '职级']:
                    if key not in emp_idx:
                        return self.send_json({'message': '员工配置缺少必要列'}, 400)
                for key in ['职级', '工作行为']:
                    if key not in beh_idx:
                        return self.send_json({'message': '行为配置缺少必要列'}, 400)

                employees = []
                for i, row in enumerate(emp_rows[1:], start=1):
                    dept = row[emp_idx['部门']] if emp_idx['部门'] < len(row) else ''
                    name = row[emp_idx['员工姓名']] if emp_idx['员工姓名'] < len(row) else ''
                    level = row[emp_idx['职级']] if emp_idx['职级'] < len(row) else ''
                    if dept and name and level:
                        employees.append({'id': i, 'department': str(dept).strip(), 'name': str(name).strip(), 'level': str(level).strip()})

                behaviors = {}
                for row in behavior_rows[1:]:
                    level = row[beh_idx['职级']] if beh_idx['职级'] < len(row) else ''
                    behavior = row[beh_idx['工作行为']] if beh_idx['工作行为'] < len(row) else ''
                    if level and behavior:
                        lv = str(level).strip()
                        behaviors.setdefault(lv, []).append(str(behavior).strip())

                departments = sorted(list({e['department'] for e in employees}))
                write_json(CONFIG_PATH, {'departments': departments, 'employees': employees, 'behaviorsByLevel': behaviors})
                return self.send_json({'message': 'Excel 导入成功'})
            finally:
                if temp.exists():
                    temp.unlink()

        self.send_error(404)

    def do_PUT(self):
        if self.path.startswith('/api/admin/') and not self.auth_ok():
            return self.send_json({'message': '未授权'}, 401)
        if self.path == '/api/admin/config':
            body = self.parse_json_body()
            if not isinstance(body.get('departments'), list) or not isinstance(body.get('employees'), list) or not isinstance(body.get('behaviorsByLevel'), dict):
                return self.send_json({'message': '配置格式不正确'}, 400)
            write_json(CONFIG_PATH, body)
            return self.send_json({'message': '配置已更新'})
        self.send_error(404)


if __name__ == '__main__':
    server = ThreadingHTTPServer(('0.0.0.0', int(os.environ.get('PORT', '3000'))), Handler)
    print('server running at http://localhost:3000')
    server.serve_forever()
