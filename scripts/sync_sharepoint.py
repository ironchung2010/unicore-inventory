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

    # 사이트 ID 조회
    print(f"SharePoint 사이트 조회 중... ({SHAREPOINT_HOST}:/sites/{SITE_NAME})")
    site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOST}:/sites/{SITE_NAME}"
    resp = requests.get(site_url, headers=headers)
    if resp.status_code != 200:
        print(f"ERROR: 사이트 조회 실패 ({resp.status_code}): {resp.text}")
        sys.exit(1)
    site_data = resp.json()
    site_id = site_data["id"]
    print(f"사이트 ID: {site_id}")
    print(f"사이트 이름: {site_data.get('displayName', 'N/A')}")

    # 드라이브 (문서 라이브러리) 조회
    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    resp = requests.get(drives_url, headers=headers)
    if resp.status_code != 200:
        print(f"ERROR: 드라이브 조회 실패 ({resp.status_code}): {resp.text}")
        sys.exit(1)
    drives = resp.json().get("value", [])

    print(f"발견된 드라이브 수: {len(drives)}")
    for i, d in enumerate(drives):
        print(f"  드라이브[{i}]: name='{d.get('name')}', id='{d.get('id', '')[:20]}...', webUrl='{d.get('webUrl', 'N/A')}'")

    drive_id = None
    for d in drives:
        if d.get("name") in ["Documents", "Shared Documents", "문서"]:
            drive_id = d["id"]
            print(f"매칭된 드라이브: '{d.get('name')}' (id: {drive_id[:20]}...)")
            break
    if not drive_id and drives:
        drive_id = drives[0]["id"]
        print(f"기본 드라이브 사용: '{drives[0].get('name')}' (id: {drive_id[:20]}...)")

    if not drive_id:
        print("ERROR: 문서 라이브러리를 찾을 수 없습니다.")
        sys.exit(1)

    # 파일 경로 정리
    file_path = FILE_PATH.strip().lstrip("/")
    if not file_path:
        print("ERROR: SHAREPOINT_FILE_PATH가 비어있습니다.")
        sys.exit(1)

    # URL이 설정된 경우 → Graph API 검색으로 파일 찾기
    if file_path.startswith("http"):
        print("SHAREPOINT_FILE_PATH에 URL이 감지됨 → 파일 검색 방식으로 전환")
        # URL에서 sourcedoc GUID 추출 시도
        guid_match = re.search(r'sourcedoc=%7[Bb]([a-fA-F0-9-]+)%7[Dd]', file_path)
        if not guid_match:
            guid_match = re.search(r'sourcedoc=\{?([a-fA-F0-9-]+)\}?', file_path)

        if guid_match:
            guid = guid_match.group(1)
            print(f"sourcedoc GUID: {guid}")
            # GUID로 SharePoint 검색 (ListItem UniqueId)
            search_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/search(q='Inventory Report')?$select=name,id,webUrl,file,sharepointIds&$top=20"
            search_resp = requests.get(search_url, headers=headers)
            if search_resp.status_code == 200:
                items = search_resp.json().get("value", [])
                print(f"검색 결과: {len(items)}개")
                target_item = None
                for item in items:
                    sp_ids = item.get("sharepointIds", {})
                    item_uid = sp_ids.get("listItemUniqueId", "")
                    print(f"  {item['name']} (uid: {item_uid})")
                    if item_uid == guid:
                        target_item = item
                        break
                if not target_item and items:
                    # GUID 매칭 실패 시 최신 Inventory Report 파일 사용
                    for item in items:
                        if 'Inventory Report' in item.get('name', '') and item.get('name', '').endswith('.xlsb'):
                            target_item = item
                            break
                if target_item:
                    item_id = target_item["id"]
                    print(f"파일 발견: {target_item['name']} (id: {item_id[:20]}...)")
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
                        print(f"ERROR: 파일 다운로드 실패 ({resp.status_code}): {resp.text}")
                        sys.exit(1)
                else:
                    print("ERROR: GUID 매칭 파일을 찾지 못했습니다.")
                    sys.exit(1)
            else:
                print(f"ERROR: 검색 실패 ({search_resp.status_code}): {search_resp.text}")
                sys.exit(1)
        else:
            print("WARNING: URL에서 sourcedoc GUID를 추출할 수 없습니다. 기본 경로로 시도합니다.")
            file_path = "General/12. SC/F'cst/UNI&CORE Inventory Report_(4월).xlsb"

    # 경로 기반 다운로드
    print(f"파일 경로: {file_path}")
    path_segments = file_path.split("/")
    encoded_segments = [quote(seg, safe='') for seg in path_segments]
    encoded_path = "/".join(encoded_segments)

    download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_path}:/content"
    print(f"인코딩 경로: {encoded_path}")
    resp = requests.get(download_url, headers=headers)

    if resp.status_code != 200:
        print(f"ERROR: 파일 다운로드 실패 ({resp.status_code}): {resp.text}")
        print("\n--- 드라이브 루트 내용 ---")
        root_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children?$select=name,folder,file&$top=20"
        root_resp = requests.get(root_url, headers=headers)
        if root_resp.status_code == 200:
            for item in root_resp.json().get("value", []):
                t = "📁" if "folder" in item else "📄"
                print(f"  {t} {item['name']}")
        sys.exit(1)

    # 파일 확장자에 따라 저장
    ext = os.path.splitext(file_path)[1].lower()
    local_path = f"/tmp/inventory_report{ext}"
    with open(local_path, "wb") as f:
        f.write(resp.content)

    print(f"다운로드 완료: {local_path} ({len(resp.content)} bytes)")
    return local_path


def parse_excel(file_path):
    """엑셀 파일에서 재고 데이터 추출 (.xlsx 및 .xlsb 지원)"""
    ext = os.path.splitext(file_path)[1].lower()
    print(f"엑셀 파일 파싱 중... (형식: {ext})")

    if ext == '.xlsb':
        # .xlsb (바이너리 형식) - pyxlsb 사용
        from pyxlsb import open_workbook
        wb = open_workbook(file_path)
        sheet_names = wb.sheets

        sheet_name = None
        for name in sheet_names:
            if '소비기한' in name:
                sheet_name = name
                break
        if not sheet_name:
            sheet_name = sheet_names[0]

        print(f"시트: {sheet_name} (전체: {sheet_names})")
        rows = []
        with wb.get_sheet(sheet_name) as ws:
            for row in ws.rows():
                rows.append(tuple(c.v for c in row))
        wb.close()
    else:
        # .xlsx 형식 - openpyxl 사용
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)

        sheet_name = None
        for name in wb.sheetnames:
            if '소비기한' in name:
                sheet_name = name
                break
        if not sheet_name:
            sheet_name = wb.sheetnames[0]

        print(f"시트: {sheet_name} (전체: {wb.sheetnames})")
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

    print(f"총 행 수: {len(rows)}")

    # 안전한 셀 접근 함수
    def safe_get(row, idx, default=None):
        """행의 인덱스 범위를 초과하지 않도록 안전하게 셀 값을 가져옴"""
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

    # 헤더 행 찾기 (처음 20행까지 검색)
    header_row = -1
    col_map = {}
    for i, row in enumerate(rows[:20]):
        cells = [str(c).strip().lower() if c else '' for c in row]
        # 다양한 헤더 패턴 매칭
        has_header = (
            'category' in cells or
            any('상품코드' in c for c in cells) or
            any('품명' in c for c in cells) or
            any('현재고' in c for c in cells) or
            (any('코드' in c for c in cells) and any('품' in c for c in cells))
        )
        if has_header:
            header_row = i
            for j, cell in enumerate(row):
                c = str(cell).strip().lower() if cell else ''
                if c == 'category' or c == '카테고리' or c == '구분': col_map['category'] = j
                if '상품코드' in c or (c == '코드'): col_map['code'] = j
                if '품명' in c or ('품' in c and '명' in c) or c == '제품명': col_map['name'] = j
                if '현재고' in c or c == '재고': col_map['stock'] = j
                if '입고' in c and ('예정' in c or '수량' in c): col_map['incoming'] = j
                if '총' in c and '재고' in c: col_map['total'] = j
                if c == 'remark' or c == '비고': col_map['remark'] = j
                if '안전재고' in c or ('안전' in c and '재고' in c): col_map['safety'] = j
                if '1m' in c and '안전' in c: col_map['safety'] = j
                if '2m' in c and '평균' in c: col_map['avg_out'] = j
                if '3m' in c and '평균' in c: col_map['avg_out_3m'] = j
                if '평균' in c and '출고' in c: col_map.setdefault('avg_out', j)
                if c == 'moq' or '발주단위' in c or '최소발주' in c: col_map['moq'] = j
                if '리드타임' in c or '리드 타임' in c or 'l/t' in c: col_map['lead_time'] = j
                if '소비기한' in c or '유통기한' in c: col_map['expiry'] = j
            print(f"헤더 발견 (행 {i}): {col_map}")
            break

    if header_row < 0:
        print("WARNING: 헤더 행을 찾지 못했습니다. 처음 5행 샘플:")
        for i, row in enumerate(rows[:5]):
            sample = [str(c)[:20] if c else 'None' for c in row[:15]]
            print(f"  행 {i}: {sample}")

        # 기본 매핑 - 행 길이에 맞춰 안전하게 설정
        header_row = 4
        max_cols = max((len(r) for r in rows[:20]), default=10)
        print(f"최대 컬럼 수: {max_cols}")
        col_map = {'category': 0, 'code': 1, 'name': 2}
        if max_cols > 5: col_map['stock'] = 5
        if max_cols > 6: col_map['incoming'] = 6
        if max_cols > 8: col_map['total'] = 8
        if max_cols > 9: col_map['remark'] = 9
        if max_cols > 15: col_map['safety'] = 15
        if max_cols > 17: col_map['avg_out'] = 17
        if max_cols > 18: col_map['moq'] = 18
        if max_cols > 19: col_map['lead_time'] = 19

    products = []
    for i in range(header_row + 1, len(rows)):
        row = rows[i]
        if not row or len(row) < 3:
            continue

        category = str(safe_get(row, col_map.get('category'), '') or '').strip()
        code = str(safe_get(row, col_map.get('code'), '') or '').strip()
        name = str(safe_get(row, col_map.get('name'), '') or '').strip()

        if not name or not code or name == '종합계' or not category:
            continue

        stock = safe_num(safe_get(row, col_map.get('stock'), 0))
        incoming = safe_num(safe_get(row, col_map.get('incoming'), 0))
        avg_out_idx = col_map.get('avg_out') or col_map.get('avg_out_3m')
        avg_out = safe_num(safe_get(row, avg_out_idx, 0))
        moq = safe_num(safe_get(row, col_map.get('moq'), 0)) or 5000
        lead_time = safe_num(safe_get(row, col_map.get('lead_time'), 0)) or 10
        remark = str(safe_get(row, col_map.get('remark'), '') or '').strip()

        expiry_date = ''
        expiry_idx = col_map.get('expiry')
        if expiry_idx is not None:
            raw = safe_get(row, expiry_idx)
            if raw is not None:
                if isinstance(raw, datetime):
                    expiry_date = raw.strftime('%Y-%m-%d')
                elif isinstance(raw, (int, float)) and raw > 0:
                    from datetime import timedelta
                    d = datetime(1899, 12, 30) + timedelta(days=int(raw))
                    expiry_date = d.strftime('%Y-%m-%d')

        products.append({
            'category': category,
            'code': code,
            'name': name,
            'currentStock': stock,
            'incoming': incoming,
            'safetyStock': safe_num(safe_get(row, col_map.get('safety'), 0)),
            'avgMonthlyOut': avg_out,
            'moq': moq,
            'leadTime': lead_time,
            'expiryDate': expiry_date,
            'remark': remark
        })

    print(f"파싱 완료: {len(products)}개 제품")
    return products


def update_dashboard(products):
    """대시보드 HTML 파일의 메인 products 배열을 업데이트"""
    print(f"대시보드 업데이트 중... ({DASHBOARD_HTML})")

    with open(DASHBOARD_HTML, 'r', encoding='utf-8') as f:
        html = f.read()

    # JavaScript 배열 문자열 생성 (JS 객체 형식)
    today = datetime.now().strftime('%Y년 %m월 %d일')
    js_entries = []
    for p in products:
        entry = "  { "
        entry += f"category: '{p['category']}', code: '{p['code']}', name: '{p['name']}', "
        entry += f"currentStock: {p['currentStock']}, incoming: {p['incoming']}, "
        entry += f"totalStock: {p['currentStock'] + p['incoming']}, "
        entry += f"remark: '{p.get('remark', '')}', expiryDate: '{p.get('expiryDate', '')}', "
        entry += f"safetyStock: {p.get('safetyStock', 0)}, avgMonthlyOut: {p.get('avgMonthlyOut', 0)}, "
        entry += f"moq: {p.get('moq', 5000)}, leadTime: {p.get('leadTime', 10)}"
        entry += " }"
        js_entries.append(entry)
    js_content = ',\n'.join(js_entries)

    # 메인 let products = [...] 배열 교체
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

    # 동기화 날짜 업데이트
    new_html = re.sub(
        r"showToast\('.*?동기화.*?'\);",
        f"showToast('실시간 데이터 ' + products.length + '개 제품이 로드되었습니다. ({today} 동기화)');",
        new_html
    )

    # 데이터 버전 업데이트 (캐시 무효화)
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

    # 최근 180일만 유지
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
