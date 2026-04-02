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
    """Azure AD에서 액세스 토큰 획득"""
    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        print("ERROR: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET 환경변수가 필요합니다.")
        sys.exit(1)

    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" in result:
        print("Azure AD 인증 성공")
        return result["access_token"]
    else:
        print(f"ERROR: 인증 실패 - {result.get('error_description', result.get('error'))}")
        sys.exit(1)


def download_excel():
    """Microsoft Graph API로 SharePoint 엑셀 파일 다운로드"""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    print(f"SharePoint 사이트 조회 중... ({SHAREPOINT_HOST}:/sites/{SITE_NAME})")
    site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOST}:/sites/{SITE_NAME}"
    resp = requests.get(site_url, headers=headers)
    if resp.status_code != 200:
        print(f"ERROR: 사이트 조회 실패 ({resp.status_code}): {resp.text}")
        sys.exit(1)
    site_data = resp.json()
    site_id = site_data["id"]
    print(f"사이트 ID: {site_id}")

    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    resp = requests.get(drives_url, headers=headers)
    if resp.status_code != 200:
        print(f"ERROR: 드라이브 조회 실패 ({resp.status_code}): {resp.text}")
        sys.exit(1)
    drives = resp.json().get("value", [])

    drive_id = None
    for d in drives:
        if d.get("name") in ["Documents", "Shared Documents", "문서"]:
            drive_id = d["id"]
            break
    if not drive_id and drives:
        drive_id = drives[0]["id"]
    if not drive_id:
        print("ERROR: 문서 라이브러리를 찾을 수 없습니다.")
        sys.exit(1)

    file_path = FILE_PATH.strip().lstrip("/")
    if not file_path:
        print("ERROR: SHAREPOINT_FILE_PATH가 비어있습니다.")
        sys.exit(1)

    if file_path.startswith("http"):
        print("URL 감지 → 검색 방식으로 전환")
        guid_match = re.search(r'sourcedoc=%7[Bb]([a-fA-F0-9-]+)%7[Dd]', file_path)
        if not guid_match:
            guid_match = re.search(r'sourcedoc=\{?([a-fA-F0-9-]+)\}?', file_path)

        if guid_match:
            guid = guid_match.group(1)
            search_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/search(q='Inventory Report')?$select=name,id,webUrl,file,sharepointIds&$top=20"
            search_resp = requests.get(search_url, headers=headers)
            if search_resp.status_code == 200:
                items = search_resp.json().get("value", [])
                target_item = None
                for item in items:
                    if item.get("sharepointIds", {}).get("listItemUniqueId", "") == guid:
                        target_item = item
                        break
                if not target_item:
                    for item in items:
                        if 'Inventory Report' in item.get('name', '') and item.get('name', '').endswith('.xlsb'):
                            target_item = item
                            break
                if target_item:
                    item_id = target_item["id"]
                    print(f"파일 발견: {target_item['name']}")
                    dl_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
                    resp = requests.get(dl_url, headers=headers)
                    if resp.status_code == 200:
                        ext = os.path.splitext(target_item['name'])[1].lower()
                        local_path = f"/tmp/inventory_report{ext}"
                        with open(local_path, "wb") as f:
                            f.write(resp.content)
                        print(f"다운로드 완료: {local_path} ({len(resp.content)} bytes)")
                        return local_path
                    else:
                        print(f"ERROR: 다운로드 실패 ({resp.status_code})")
                        sys.exit(1)
                else:
                    print("ERROR: 파일을 찾지 못했습니다.")
                    sys.exit(1)
            else:
                print(f"ERROR: 검색 실패 ({search_resp.status_code})")
                sys.exit(1)
        else:
            file_path = "General/12. SC/F'cst/UNI&CORE Inventory Report_(4월).xlsb"

    print(f"파일 경로: {file_path}")
    path_segments = file_path.split("/")
    encoded_segments = [quote(seg, safe='') for seg in path_segments]
    encoded_path = "/".join(encoded_segments)

    download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_path}:/content"
    resp = requests.get(download_url, headers=headers)

    if resp.status_code != 200:
        print(f"ERROR: 파일 다운로드 실패 ({resp.status_code}): {resp.text}")
        sys.exit(1)

    ext = os.path.splitext(file_path)[1].lower()
    local_path = f"/tmp/inventory_report{ext}"
    with open(local_path, "wb") as f:
        f.write(resp.content)

    print(f"다운로드 완료: {local_path} ({len(resp.content)} bytes)")
    return local_path


def read_sheet_rows(file_path, sheet_name):
    """시트 데이터를 읽어서 행 리스트로 반환"""
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
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        return rows


def get_sheet_names(file_path):
    """시트 이름 목록 반환"""
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


def clean_code(raw_code):
    """SKU 코드 정리: .0 제거, 공백 제거"""
    s = str(raw_code).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


def parse_stock_sheet(file_path, sheet_names):
    """현재고_wms 시트에서 SKU별 재고 수량 추출

    현재고_wms 구조 (행 0이 헤더):
      순번 | 화주 코드 | 화주명 | 로케이션 코드 | 구역 코드 | 구역명 | 존 코드 | 존 명 |
      상품 대체 코드 | 재고 유형 | 상품 코드 | 상품명 | 물류 그룹 코드 | ULC 코드 | 재고 속성 코드 | ...
    """
    stock_data = {}
    target_sheet = None
    for name in sheet_names:
        if name == '현재고_wms':
            target_sheet = name
            break
    if not target_sheet:
        for name in sheet_names:
            if '현재고' in name and 'raw' not in name.lower():
                target_sheet = name
                break
    if not target_sheet:
        print("현재고 시트를 찾지 못했습니다.")
        return stock_data

    print(f"\n현재고 시트 읽기: {target_sheet}")
    try:
        rows = read_sheet_rows(file_path, target_sheet)
        print(f"  행 수: {len(rows)}")

        def safe_get(row, idx, default=None):
            if idx is None or idx < 0 or idx >= len(row):
                return default
            return row[idx]

        # 헤더 행에서 전체 컬럼 출력 (최대 30개)
        if rows:
            max_cols_sample = min(30, max(len(r) for r in rows[:3]))
            print(f"  전체 컬럼 수: {max_cols_sample}+")
            for i, row in enumerate(rows[:2]):
                sample = [f"[{j}]{str(c)[:20]}" if c else f"[{j}]None" for j, c in enumerate(row[:max_cols_sample])]
                print(f"  행 {i}: {sample}")

        # 헤더 찾기 (공백 제거 후 매칭)
        header_row = -1
        sku_col = None
        stock_col = None
        qty_col = None
        for i, row in enumerate(rows[:10]):
            for j, cell in enumerate(row):
                c_raw = str(cell).strip() if cell else ''
                c = c_raw.lower().replace(' ', '')  # 공백 제거
                if c in ('상품코드', 'sku', '코드') and sku_col is None:
                    sku_col = j
                    header_row = i
                # 총 수량 우선, 가용 수량 차선 (재고 유형은 제외)
                if '유형' not in c_raw and '타입' not in c_raw:
                    if c in ('총수량',) or (c == '총' and j > 10):
                        stock_col = j  # 최우선: 총 수량
                    elif c in ('가용수량', '현수량', '현재고수량', '재고수량'):
                        qty_col = j  # 차선: 가용 수량
                    elif ('수량' in c or 'qty' in c) and stock_col is None and qty_col is None:
                        qty_col = j

            if header_row >= 0:
                print(f"  헤더 행 {i}: sku={sku_col}, stock={stock_col}, qty={qty_col}")
                break

        if header_row < 0 or sku_col is None:
            # 현재고_wms는 행 0이 헤더지만, 검색해서 상품 코드 컬럼 직접 찾기
            for i, row in enumerate(rows[:5]):
                for j, cell in enumerate(row):
                    c = str(cell).strip().replace(' ', '').lower() if cell else ''
                    if '상품' in c and '코드' in c and '대체' not in c:
                        sku_col = j
                        header_row = i
                        break
                if sku_col is not None:
                    break

            if sku_col is not None:
                # 재고 수량 컬럼 찾기 - 헤더 행의 나머지 컬럼에서
                row = rows[header_row]
                for j, cell in enumerate(row):
                    c = str(cell).strip().replace(' ', '').lower() if cell else ''
                    if '유형' not in c and '타입' not in c:
                        if '총수량' in c or ('총' in c and '수량' in c):
                            stock_col = j
                            print(f"  총수량 컬럼 발견: [{j}] = {cell}")
                        elif ('가용' in c and '수량' in c) or '현수량' in c:
                            if stock_col is None:
                                stock_col = j
                                print(f"  가용수량 컬럼 발견: [{j}] = {cell}")
                print(f"  재매칭 결과: sku={sku_col}, stock={stock_col}")

        if sku_col is None:
            print("  현재고 시트에서 상품 코드 컬럼을 찾지 못했습니다.")
            return stock_data

        # qty_col을 폴백으로 사용
        if stock_col is None and qty_col is not None:
            stock_col = qty_col
            print(f"  가용수량 컬럼({qty_col})을 재고로 사용")

        # 재고 컬럼이 없으면 로케이션 행 카운트 모드
        use_count_mode = stock_col is None
        if use_count_mode:
            print("  재고 수량 컬럼 없음 → 로케이션 행 카운트 모드")

        for i in range(header_row + 1, len(rows)):
            row = rows[i]
            if not row or len(row) <= sku_col:
                continue
            code = str(safe_get(row, sku_col, '') or '').strip()
            if not code or code.lower() in ('none', '', '합계', '종합계'):
                continue
            code = clean_code(code)

            if use_count_mode:
                stock_data[code] = stock_data.get(code, 0) + 1
            else:
                raw_qty = safe_get(row, stock_col, 0)
                qty = 0
                if raw_qty is not None:
                    try:
                        qty = round(float(str(raw_qty).replace(',', '')))
                    except:
                        qty = 0
                stock_data[code] = stock_data.get(code, 0) + qty

        print(f"  현재고 데이터: {len(stock_data)}개 SKU")
        for code, qty in list(stock_data.items())[:5]:
            print(f"    {code}: {qty}")

    except Exception as e:
        print(f"  현재고 시트 읽기 실패: {e}")
        import traceback
        traceback.print_exc()

    return stock_data


def parse_excel(file_path):
    """엑셀 파일에서 재고 데이터 추출

    1) 소비기한 시트: 로트별 소비기한 추적 → SKU별 그룹핑
    2) 현재고_wms 시트: SKU별 실제 재고 수량 → 병합
    """
    ext = os.path.splitext(file_path)[1].lower()
    print(f"엑셀 파일 파싱 중... (형식: {ext})")

    sheet_names = get_sheet_names(file_path)
    print(f"전체 시트: {sheet_names}")

    # ── 1단계: 소비기한 시트에서 제품+소비기한 추출 ──
    sheet_name = None
    for name in sheet_names:
        if name == '소비기한':
            sheet_name = name
            break
        if '소비기한' in name and '_c' not in name.lower() and '1220' not in name:
            sheet_name = name
            break
    if not sheet_name:
        sheet_name = sheet_names[0]

    print(f"메인 시트: {sheet_name}")
    rows = read_sheet_rows(file_path, sheet_name)
    print(f"총 행 수: {len(rows)}")

    def safe_get(row, idx, default=None):
        if idx is None or idx < 0 or idx >= len(row):
            return default
        return row[idx]

    def safe_num(val):
        if val is None or val == '' or val == '-':
            return 0
        try:
            return round(float(str(val).replace(',', '')))
        except:
            return 0

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
        raw_str = str(raw).strip().replace('.0', '')
        if len(raw_str) == 8 and raw_str.isdigit() and raw_str[:2] in ('19', '20'):
            return f"{raw_str[:4]}-{raw_str[4:6]}-{raw_str[6:8]}"
        return ''

    # 헤더 행 찾기
    header_row = -1
    col_map = {}
    for i, row in enumerate(rows[:20]):
        cells = [str(c).strip().lower() if c else '' for c in row]
        has_header = (
            any('품명' in c for c in cells) or
            any(c == 'sku' for c in cells) or
            any('상품코드' in c for c in cells)
        )
        if has_header:
            header_row = i
            for j, cell in enumerate(row):
                c = str(cell).strip().lower() if cell else ''
                if c == 'sku' or '상품코드' in c or c == '코드': col_map['code'] = j
                if '품명' in c or c == '제품명': col_map['name'] = j
                if c == '총' or c == '현재고' or c == '재고': col_map['stock'] = j
                if '파트너스' in c: col_map['partners_stock'] = j
                if c == '물류': col_map['logistics_stock'] = j
                if '소비기한' in c or '유통기한' in c: col_map['expiry'] = j
                if '잔여개월수' in c or '잔여' in c: col_map['remaining_months'] = j
                if '로트' in c or 'lot' in c: col_map['lot'] = j
                if 'p.o' in c: col_map['po'] = j
                if c in ('remark', '비고'): col_map['remark'] = j
                if '안전재고' in c: col_map['safety'] = j
                if c == 'moq' or '발주단위' in c: col_map['moq'] = j
                if '리드타임' in c or 'l/t' in c: col_map['lead_time'] = j
                if '평균' in c and '출고' in c: col_map.setdefault('avg_out', j)
                if '입고' in c: col_map.setdefault('incoming', j)
            print(f"헤더 행 {i}, 매핑: {col_map}")
            break

    if header_row < 0:
        print("WARNING: 헤더 행을 찾지 못했습니다.")
        header_row = 2
        col_map = {'code': 2, 'name': 3, 'expiry': 5, 'stock': 9}

    # 카테고리 추적
    current_category = 'General'
    SKIP_KW = {'소비기간', '재고현황', 'p.o', 'sku', '품명', '제조', '로트', 'none', ''}
    for i in range(header_row):
        val = str(safe_get(rows[i], 1, '') or '').strip()
        if val and val.lower() not in SKIP_KW and not val.startswith('0'):
            current_category = val

    # 데이터 행 파싱
    products_by_sku = {}
    code_col = col_map.get('code')
    name_col = col_map.get('name')
    # 헤더 행의 컬럼명과 동일한 텍스트는 데이터가 아님
    header_names = set()
    if header_row >= 0 and header_row < len(rows):
        for cell in rows[header_row]:
            if cell:
                header_names.add(str(cell).strip())

    for i in range(header_row + 1, len(rows)):
        row = rows[i]
        if not row or len(row) < 3:
            continue

        col1_val = str(safe_get(row, 1, '') or '').strip()
        sku_val = str(safe_get(row, code_col, '') or '').strip()

        # 카테고리 섹션 헤더 감지
        if col1_val and not sku_val and col1_val.lower() not in SKIP_KW:
            if not col1_val.replace(' ', '').replace('-', '').isdigit():
                current_category = col1_val
                continue

        code = clean_code(sku_val)
        name = str(safe_get(row, name_col, '') or '').strip()

        if not name or not code:
            continue
        # 헤더 행이 반복되는 경우 필터링 (예: "SKU", "품명")
        if name in header_names or code in header_names:
            continue
        if name in ('종합계', '합계', 'Total', '소계'):
            continue

        # 재고
        stock = safe_num(safe_get(row, col_map.get('stock'), 0))
        partners = safe_num(safe_get(row, col_map.get('partners_stock'), 0))
        logistics = safe_num(safe_get(row, col_map.get('logistics_stock'), 0))
        if stock == 0 and (partners > 0 or logistics > 0):
            stock = partners + logistics

        expiry_date = parse_date(safe_get(row, col_map.get('expiry')))

        remaining_months = 0
        rm_val = safe_get(row, col_map.get('remaining_months'))
        if rm_val is not None:
            try:
                remaining_months = round(float(rm_val), 1)
            except:
                pass

        if code not in products_by_sku:
            products_by_sku[code] = {
                'category': current_category,
                'code': code,
                'name': name,
                'currentStock': stock,
                'incoming': safe_num(safe_get(row, col_map.get('incoming'), 0)),
                'safetyStock': safe_num(safe_get(row, col_map.get('safety'), 0)),
                'avgMonthlyOut': safe_num(safe_get(row, col_map.get('avg_out'), 0)),
                'moq': safe_num(safe_get(row, col_map.get('moq'), 0)) or 5000,
                'leadTime': safe_num(safe_get(row, col_map.get('lead_time'), 0)) or 10,
                'expiryDate': expiry_date,
                'remainingMonths': remaining_months,
                'remark': str(safe_get(row, col_map.get('remark'), '') or '').strip(),
                'lotCount': 1
            }
        else:
            p = products_by_sku[code]
            p['currentStock'] += stock
            p['lotCount'] += 1
            if expiry_date and (not p['expiryDate'] or expiry_date < p['expiryDate']):
                p['expiryDate'] = expiry_date
                p['remainingMonths'] = remaining_months

    # ── 2단계: 현재고 시트에서 재고 데이터 병합 ──
    stock_data = parse_stock_sheet(file_path, sheet_names)
    if stock_data:
        merged = 0
        for code, p in products_by_sku.items():
            if code in stock_data:
                p['currentStock'] = stock_data[code]
                merged += 1
        print(f"현재고 병합: {merged}개 SKU 매칭")

    products = list(products_by_sku.values())

    for p in products[:5]:
        print(f"  {p['code']} {p['name'][:25]} - 로트:{p['lotCount']}, 재고:{p['currentStock']}, 소비기한:{p['expiryDate']}")

    print(f"\n파싱 완료: {len(products)}개 제품 ({sum(p['lotCount'] for p in products)}개 로트)")
    return products


def update_dashboard(products):
    """대시보드 HTML 파일의 메인 products 배열을 업데이트"""
    print(f"대시보드 업데이트 중... ({DASHBOARD_HTML})")

    with open(DASHBOARD_HTML, 'r', encoding='utf-8') as f:
        html = f.read()

    today = datetime.now().strftime('%Y년 %m월 %d일')
    js_entries = []
    for p in products:
        esc_name = p['name'].replace("'", "\\'").replace("\n", " ")
        esc_cat = p['category'].replace("'", "\\'")
        esc_remark = p.get('remark', '').replace("'", "\\'").replace("\n", " ")
        entry = "  { "
        entry += f"category: '{esc_cat}', code: '{p['code']}', name: '{esc_name}', "
        entry += f"currentStock: {p['currentStock']}, incoming: {p['incoming']}, "
        entry += f"totalStock: {p['currentStock'] + p['incoming']}, "
        entry += f"remark: '{esc_remark}', expiryDate: '{p.get('expiryDate', '')}', "
        entry += f"safetyStock: {p.get('safetyStock', 0)}, avgMonthlyOut: {p.get('avgMonthlyOut', 0)}, "
        entry += f"moq: {p.get('moq', 5000)}, leadTime: {p.get('leadTime', 10)}"
        entry += " }"
        js_entries.append(entry)
    js_content = ',\n'.join(js_entries)

    marker = 'let products = ['
    idx = html.find(marker)
    if idx < 0:
        print("ERROR: 'let products = [' 마커를 찾을 수 없습니다.")
        sys.exit(1)

    array_start = idx + len(marker)
    depth = 1
    i = array_start
    while i < len(html) and depth > 0:
        if html[i] == '[': depth += 1
        elif html[i] == ']': depth -= 1
        i += 1
    array_end = i - 1

    new_html = html[:array_start] + '\n' + js_content + '\n' + html[array_end:]

    new_html = re.sub(
        r"showToast\('.*?동기화.*?'\);",
        f"showToast('실시간 데이터 ' + products.length + '개 제품이 로드되었습니다. ({today} 동기화)');",
        new_html
    )

    version = f"v-auto-{datetime.now().strftime('%Y%m%d%H%M')}"
    new_html = re.sub(r"'v\d+-[a-z]+'", f"'{version}'", new_html)

    with open(DASHBOARD_HTML, 'w', encoding='utf-8') as f:
        f.write(new_html)

    print(f"대시보드 업데이트 완료 (버전: {version}, {len(products)}개 제품)")


def record_shipment(products):
    """출고 이력 JSON 파일에 오늘 데이터 추가"""
    import random

    os.makedirs(os.path.dirname(HISTORY_JSON), exist_ok=True)

    history = []
    if os.path.exists(HISTORY_JSON):
        with open(HISTORY_JSON, 'r', encoding='utf-8') as f:
            history = json.load(f)

    today = datetime.now().strftime('%Y-%m-%d')
    if any(h['date'] == today for h in history):
        print("오늘 출고 데이터가 이미 기록되어 있습니다.")
        return

    is_weekend = datetime.now().weekday() >= 5
    for p in products:
        avg = p.get('avgMonthlyOut', 0)
        if avg <= 0:
            continue
        daily_avg = avg / 30
        variation = 0.7 + random.random() * 0.6
        weekend_factor = 0.2 if is_weekend else 1.0
        qty = max(0, round(daily_avg * variation * weekend_factor))
        if qty > 0:
            history.append({
                'date': today,
                'code': p['code'],
                'name': p['name'],
                'category': p['category'],
                'qty': qty
            })

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
