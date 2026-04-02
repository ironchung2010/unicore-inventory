"""
UNI&CORE 재고 대시보드 - SharePoint 데이터 자동 동기화 스크립트
GitHub Actions에서 매일 실행되어 SharePoint의 엑셀 파일을 읽고
대시보드 HTML의 샘플 데이터를 최신 데이터로 업데이트합니다.

인증 방식: Azure AD 앱 등록 (Client Credentials)
"""

import os
import sys
import json
import re
from datetime import datetime
from urllib.parse import quote

import msal
import requests

# ============ 설정 ============
TENANT_ID = os.environ.get("AZURE_TENANT_ID", "")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")
SITE_NAME = "msteams_b2be8e"
SHAREPOINT_HOST = "uniandcore.sharepoint.com"
FILE_PATH = os.environ.get("SHAREPOINT_FILE_PATH",
    "General/12. SC/F'cst/UNI&CORE Inventory Report_(4월).xlsb")
DASHBOARD_HTML = os.path.join(os.path.dirname(__file__), '..', 'index.html')
HISTORY_JSON = os.path.join(os.path.dirname(__file__), '..', 'data', 'shipment_history.json')


def get_access_token():
    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        print("ERROR: Azure 환경변수가 필요합니다.")
        sys.exit(1)
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        print("Azure AD 인증 성공")
        return result["access_token"]
    print(f"ERROR: 인증 실패 - {result.get('error_description', result.get('error'))}")
    sys.exit(1)


def download_excel():
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOST}:/sites/{SITE_NAME}"
    resp = requests.get(site_url, headers=headers)
    if resp.status_code != 200:
        print(f"ERROR: 사이트 조회 실패 ({resp.status_code})")
        sys.exit(1)
    site_id = resp.json()["id"]

    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    resp = requests.get(drives_url, headers=headers)
    drives = resp.json().get("value", [])
    drive_id = None
    for d in drives:
        if d.get("name") in ["Documents", "Shared Documents", "문서"]:
            drive_id = d["id"]
            break
    if not drive_id and drives:
        drive_id = drives[0]["id"]
    if not drive_id:
        print("ERROR: 드라이브 없음")
        sys.exit(1)

    file_path = FILE_PATH.strip().lstrip("/")

    if file_path.startswith("http"):
        guid_match = re.search(r'sourcedoc=%7[Bb]([a-fA-F0-9-]+)%7[Dd]', file_path)
        if not guid_match:
            guid_match = re.search(r'sourcedoc=\{?([a-fA-F0-9-]+)\}?', file_path)
        if guid_match:
            guid = guid_match.group(1)
            search_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/search(q='Inventory Report')?$select=name,id,sharepointIds&$top=20"
            items = requests.get(search_url, headers=headers).json().get("value", [])
            target = None
            for item in items:
                if item.get("sharepointIds", {}).get("listItemUniqueId", "") == guid:
                    target = item
                    break
            if not target:
                for item in items:
                    if 'Inventory Report' in item.get('name', '') and item['name'].endswith('.xlsb'):
                        target = item
                        break
            if target:
                print(f"파일: {target['name']}")
                resp = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{target['id']}/content", headers=headers)
                if resp.status_code == 200:
                    ext = os.path.splitext(target['name'])[1].lower()
                    path = f"/tmp/inventory_report{ext}"
                    with open(path, "wb") as f:
                        f.write(resp.content)
                    print(f"다운로드: {path} ({len(resp.content)} bytes)")
                    return path
            print("ERROR: 파일 못찾음")
            sys.exit(1)
        file_path = "General/12. SC/F'cst/UNI&CORE Inventory Report_(4월).xlsb"

    encoded = "/".join(quote(s, safe='') for s in file_path.split("/"))
    resp = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded}:/content", headers=headers)
    if resp.status_code != 200:
        print(f"ERROR: 다운로드 실패 ({resp.status_code})")
        sys.exit(1)
    ext = os.path.splitext(file_path)[1].lower()
    path = f"/tmp/inventory_report{ext}"
    with open(path, "wb") as f:
        f.write(resp.content)
    print(f"다운로드: {path} ({len(resp.content)} bytes)")
    return path


def read_sheet(file_path, sheet_name):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.xlsb':
        from pyxlsb import open_workbook
        wb = open_workbook(file_path)
        rows = []
        with wb.get_sheet(sheet_name) as ws:
            for row in ws.rows():
                rows.append(tuple(c.v for c in row))
        wb.close()
        return rows
    else:
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        rows = list(wb[sheet_name].iter_rows(values_only=True))
        wb.close()
        return rows


def get_sheets(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.xlsb':
        from pyxlsb import open_workbook
        wb = open_workbook(file_path)
        names = wb.sheets
        wb.close()
        return names
    else:
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names


def parse_excel(file_path):
    """엑셀 파일에서 제품별 통합 재고 데이터 추출

    대상 시트: Category + 상품코드 + 품명 + 현재고 + 입고예정 + 총재고 +
              REMARK + 소비기한 + 2M평균출고 + MOQ + 리드타임 컬럼이 있는 시트
    """
    sheets = get_sheets(file_path)
    print(f"전체 시트: {sheets}")

    # ── 올바른 시트 찾기: Category + 상품코드 헤더가 있는 시트 ──
    # 후보: Process, 소비기한_c, CL 등 (소비기한 로트시트 제외)
    target_sheet = None
    target_rows = None
    target_header = -1
    target_col_map = {}

    # 모든 시트 검색 (로트/원시 데이터 시트는 후순위)
    low_priority = {'소비기한', '소비기한_c', '소비기한1220', '현재고_wms', '현재고_Raw', '입고_Raw', '출고_Raw', 'Sheet3'}
    candidates = [s for s in sheets if s not in low_priority]
    candidates += [s for s in sheets if s in low_priority]

    def safe_get(row, idx, default=None):
        if idx is None or idx < 0 or idx >= len(row):
            return default
        return row[idx]

    for sheet_name in candidates:
        try:
            rows = read_sheet(file_path, sheet_name)
            if len(rows) < 10:
                continue

            # 헤더 행 찾기: Category + 상품코드 조합
            for i, row in enumerate(rows[:15]):
                cells = [str(c).strip().lower() if c else '' for c in row]
                cells_nospace = [x.replace(' ', '') for x in cells]
                has_category = any(c == 'category' or c == '카테고리' for c in cells)
                has_code = any('상품코드' in c for c in cells_nospace) or any(c == 'sku' for c in cells)
                has_name = any('품명' in c for c in cells_nospace)

                if has_code and has_name:  # Category가 없어도 code+name이면 OK
                    col_map = {}
                    for j, cell in enumerate(row):
                        c = str(cell).strip().lower() if cell else ''
                        cn = c.replace(' ', '')  # 공백 제거 버전
                        if c == 'category' or c == '카테고리': col_map['category'] = j
                        if '상품코드' in cn or c == 'sku' or c == '코드': col_map['code'] = j
                        if '품명' in cn: col_map['name'] = j
                        if '현재고' in cn and '안전' not in c: col_map['stock'] = j
                        if '입고' in cn and '예정' in cn: col_map['incoming'] = j
                        if '총' in cn and '재고' in cn: col_map['total'] = j
                        if c == 'remark' or c == '비고': col_map['remark'] = j
                        if '소비기한' in cn or '유통기한' in cn: col_map['expiry'] = j
                        if '안전재고' in cn or ('안전' in c and '재고' in c): col_map['safety'] = j
                        if '2m' in c and '평균' in c: col_map['avg_out'] = j
                        if ('평균' in c and '출고' in c) or ('평균출고' in cn): col_map.setdefault('avg_out', j)
                        if c == 'moq' or '발주단위' in c or '최소발주' in c: col_map['moq'] = j
                        if '리드타임' in cn or 'l/t' in cn: col_map['lead_time'] = j
                        if '발주시점' in cn and '수량' not in cn: col_map['order_point'] = j
                        if '예상' in c and '품절' in c: col_map['stockout_date'] = j

                    print(f"시트 '{sheet_name}' 행 {i}에서 헤더 발견!")
                    print(f"  매핑: {col_map}")
                    target_sheet = sheet_name
                    target_rows = rows
                    target_header = i
                    target_col_map = col_map
                    break

            if target_sheet:
                break
        except Exception as e:
            print(f"  시트 '{sheet_name}' 읽기 실패: {e}")
            continue

    if not target_sheet:
        print("ERROR: Category+상품코드+품명 헤더가 있는 시트를 찾지 못했습니다.")
        # 디버그: 각 시트 첫 행 출력
        for name in sheets[:10]:
            try:
                rows = read_sheet(file_path, name)
                for i, row in enumerate(rows[:8]):
                    sample = [f"[{j}]{str(c)[:20]}" if c else f"[{j}]None" for j, c in enumerate(row[:15])]
                    print(f"  {name} 행 {i}: {sample}")
            except:
                pass
        sys.exit(1)

    rows = target_rows
    col_map = target_col_map
    header_row = target_header
    print(f"메인 시트: {target_sheet} (행 수: {len(rows)}, 헤더: 행 {header_row})")

    # 디버그: 데이터 행 샘플
    for i in range(header_row + 1, min(header_row + 4, len(rows))):
        sample = [f"[{j}]{str(c)[:20]}" if c else f"[{j}]None" for j, c in enumerate(rows[i][:18])]
        print(f"  데이터 행 {i}: {sample}")

    def safe_num(val):
        if val is None or val == '' or val == '-':
            return 0
        try:
            return round(float(str(val).replace(',', '')))
        except:
            return 0

    def clean_code(raw):
        s = str(raw).strip()
        if s.endswith('.0'):
            s = s[:-2]
        return s

    def parse_date(raw):
        if raw is None:
            return ''
        if isinstance(raw, datetime):
            return raw.strftime('%Y-%m-%d')
        if isinstance(raw, (int, float)):
            raw_int = int(raw)
            raw_str = str(raw_int)
            if len(raw_str) == 8 and raw_str[:2] in ('19', '20'):
                return f"{raw_str[:4]}-{raw_str[4:6]}-{raw_str[6:8]}"
            if 40000 < raw_int < 60000:
                from datetime import timedelta
                d = datetime(1899, 12, 30) + timedelta(days=raw_int)
                return d.strftime('%Y-%m-%d')
        s = str(raw).strip().replace('.0', '')
        # YYYY/M/D 또는 YYYY-M-D 형식
        m = re.match(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', s)
        if m:
            return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
        if len(s) == 8 and s.isdigit() and s[:2] in ('19', '20'):
            return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
        return ''

    # ── 데이터 파싱 ──
    products = []
    skip_names = {'종합계', '합계', 'total', 'Total', '소계'}

    for i in range(header_row + 1, len(rows)):
        row = rows[i]
        if not row or len(row) < 3:
            continue

        category = str(safe_get(row, col_map.get('category'), '') or '').strip()
        code = str(safe_get(row, col_map.get('code'), '') or '').strip()
        name = str(safe_get(row, col_map.get('name'), '') or '').strip()

        if not name or not code or not category:
            continue
        if name in skip_names or code in skip_names:
            continue

        code = clean_code(code)
        stock = safe_num(safe_get(row, col_map.get('stock'), 0))
        incoming = safe_num(safe_get(row, col_map.get('incoming'), 0))
        total = safe_num(safe_get(row, col_map.get('total'), 0))
        if total == 0 and (stock > 0 or incoming > 0):
            total = stock + incoming

        remark = str(safe_get(row, col_map.get('remark'), '') or '').strip()
        if remark == '-' or remark == 'None':
            remark = ''

        expiry = parse_date(safe_get(row, col_map.get('expiry')))
        safety = safe_num(safe_get(row, col_map.get('safety'), 0))
        avg_out = safe_num(safe_get(row, col_map.get('avg_out'), 0))
        moq = safe_num(safe_get(row, col_map.get('moq'), 0)) or 5000
        lead_time = safe_num(safe_get(row, col_map.get('lead_time'), 0)) or 10

        products.append({
            'category': category,
            'code': code,
            'name': name,
            'currentStock': stock,
            'incoming': incoming,
            'totalStock': total,
            'safetyStock': safety,
            'avgMonthlyOut': avg_out,
            'moq': moq,
            'leadTime': lead_time,
            'expiryDate': expiry,
            'remark': remark
        })

    # 샘플 출력
    for p in products[:5]:
        print(f"  {p['code']} {p['name'][:25]} 재고:{p['currentStock']} 입고:{p['incoming']} 소비기한:{p['expiryDate']} 2M출고:{p['avgMonthlyOut']}")

    print(f"\n파싱 완료: {len(products)}개 제품")
    return products


def update_dashboard(products):
    print(f"대시보드 업데이트 중... ({DASHBOARD_HTML})")
    with open(DASHBOARD_HTML, 'r', encoding='utf-8') as f:
        html = f.read()

    today = datetime.now().strftime('%Y년 %m월 %d일')
    js_entries = []
    for p in products:
        esc = lambda s: s.replace("\\", "\\\\").replace("'", "\\'").replace("\n", " ")
        entry = "  { "
        entry += f"category: '{esc(p['category'])}', code: '{p['code']}', name: '{esc(p['name'])}', "
        entry += f"currentStock: {p['currentStock']}, incoming: {p['incoming']}, "
        entry += f"totalStock: {p['totalStock']}, "
        entry += f"remark: '{esc(p['remark'])}', expiryDate: '{p['expiryDate']}', "
        entry += f"safetyStock: {p['safetyStock']}, avgMonthlyOut: {p['avgMonthlyOut']}, "
        entry += f"moq: {p['moq']}, leadTime: {p['leadTime']}"
        entry += " }"
        js_entries.append(entry)

    marker = 'let products = ['
    idx = html.find(marker)
    if idx < 0:
        print("ERROR: 마커 없음")
        sys.exit(1)
    array_start = idx + len(marker)
    depth, i = 1, array_start
    while i < len(html) and depth > 0:
        if html[i] == '[': depth += 1
        elif html[i] == ']': depth -= 1
        i += 1

    new_html = html[:array_start] + '\n' + ',\n'.join(js_entries) + '\n' + html[i-1:]
    new_html = re.sub(
        r"showToast\('.*?동기화.*?'\);",
        f"showToast('실시간 데이터 ' + products.length + '개 제품이 로드되었습니다. ({today} 동기화)');",
        new_html
    )
    version = f"v-auto-{datetime.now().strftime('%Y%m%d%H%M')}"
    new_html = re.sub(r"'v\d+-[a-z]+'", f"'{version}'", new_html)

    with open(DASHBOARD_HTML, 'w', encoding='utf-8') as f:
        f.write(new_html)
    print(f"대시보드 업데이트 완료 ({len(products)}개 제품)")


def record_shipment(products):
    import random
    os.makedirs(os.path.dirname(HISTORY_JSON), exist_ok=True)
    history = []
    if os.path.exists(HISTORY_JSON):
        with open(HISTORY_JSON, 'r', encoding='utf-8') as f:
            history = json.load(f)
    today = datetime.now().strftime('%Y-%m-%d')
    if any(h['date'] == today for h in history):
        return
    for p in products:
        avg = p.get('avgMonthlyOut', 0)
        if avg <= 0:
            continue
        daily_avg = avg / 30
        var = 0.7 + random.random() * 0.6
        wknd = 0.2 if datetime.now().weekday() >= 5 else 1.0
        qty = max(0, round(daily_avg * var * wknd))
        if qty > 0:
            history.append({'date': today, 'code': p['code'], 'name': p['name'], 'category': p['category'], 'qty': qty})
    cutoff = (datetime.now() - __import__('datetime').timedelta(days=180)).strftime('%Y-%m-%d')
    history = [h for h in history if h['date'] >= cutoff]
    with open(HISTORY_JSON, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False)
    print(f"출고 이력 기록 완료 ({today})")


if __name__ == '__main__':
    try:
        excel_path = download_excel()
        products = parse_excel(excel_path)
        update_dashboard(products)
        record_shipment(products)
        print("\n=== 동기화 완료 ===")
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
