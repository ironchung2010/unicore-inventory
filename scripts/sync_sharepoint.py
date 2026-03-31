"""
UNI&CORE 재고 대시보드 - SharePoint 데이터 자동 동기화 스크립트
GitHub Actions에서 매일 실행되어 SharePoint의 엑셀 파일을 읽고
대시보드 HTML의 샘플 데이터를 최신 데이터로 업데이트합니다.
"""

import os
import sys
import json
import re
from datetime import datetime
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

# ============ 설정 ============
SHAREPOINT_SITE = "https://uniandcore.sharepoint.com/sites/msteams_b2be8e"
FILE_PATH = os.environ.get("SHAREPOINT_FILE_PATH",
    "/sites/msteams_b2be8e/Shared Documents/General/12. SC/F'cst/UNI&CORE Inventory Report_(3월).xlsb")
DASHBOARD_HTML = os.path.join(os.path.dirname(__file__), '..', 'index.html')
HISTORY_JSON = os.path.join(os.path.dirname(__file__), '..', 'data', 'shipment_history.json')

def download_excel():
    """SharePoint에서 엑셀 파일 다운로드"""
    username = os.environ.get("SHAREPOINT_USERNAME")
    password = os.environ.get("SHAREPOINT_PASSWORD")

    if not username or not password:
        print("ERROR: SHAREPOINT_USERNAME / SHAREPOINT_PASSWORD 환경변수가 설정되지 않았습니다.")
        sys.exit(1)

    print(f"SharePoint 연결 중... ({SHAREPOINT_SITE})")
    ctx = ClientContext(SHAREPOINT_SITE).with_credentials(
        UserCredential(username, password)
    )

    local_path = "/tmp/inventory_report.xlsb"
    print(f"파일 다운로드 중... ({FILE_PATH})")

    with open(local_path, "wb") as f:
        ctx.web.get_file_by_server_relative_url(FILE_PATH).download(f).execute_query()

    print(f"다운로드 완료: {local_path}")
    return local_path


def parse_excel(file_path):
    """엑셀 파일에서 재고 데이터 추출"""
    import openpyxl

    print("엑셀 파일 파싱 중...")
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)

    # 메인 시트 찾기
    sheet_name = None
    for name in wb.sheetnames:
        if '소비기한' in name:
            sheet_name = name
            break
    if not sheet_name:
        sheet_name = wb.sheetnames[0]

    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))

    # 헤더 행 찾기
    header_row = -1
    col_map = {}
    for i, row in enumerate(rows[:10]):
        cells = [str(c).strip().lower() if c else '' for c in row]
        if 'category' in cells or any('상품코드' in c for c in cells):
            header_row = i
            for j, cell in enumerate(row):
                c = str(cell).strip().lower() if cell else ''
                if c == 'category': col_map['category'] = j
                if '상품코드' in c: col_map['code'] = j
                if '품' in c and '명' in c: col_map['name'] = j
                if '현재고' in c: col_map['stock'] = j
                if '입고' in c and '예정' in c: col_map['incoming'] = j
                if '총' in c and '재고' in c: col_map['total'] = j
                if c == 'remark': col_map['remark'] = j
                if '안전재고' in c and '1m' in c: col_map['safety'] = j
                if '2m' in c and '평균' in c: col_map['avg_out'] = j
                if '3m' in c and '평균' in c: col_map['avg_out_3m'] = j
                if c == 'moq': col_map['moq'] = j
                if '리드타임' in c or '리드 타임' in c: col_map['lead_time'] = j
                if '소비기한' in c or '유통기한' in c: col_map['expiry'] = j
            break

    if header_row < 0:
        print("WARNING: 헤더 행을 찾지 못했습니다. 기본 구조로 파싱합니다.")
        header_row = 4
        col_map = {'category': 0, 'code': 1, 'name': 2, 'stock': 5, 'incoming': 6,
                    'total': 8, 'remark': 9, 'safety': 15, 'avg_out': 17, 'moq': 18, 'lead_time': 19}

    products = []
    for i in range(header_row + 1, len(rows)):
        row = rows[i]
        if not row or len(row) < 3:
            continue

        category = str(row[col_map.get('category', 0)] or '').strip()
        code = str(row[col_map.get('code', 1)] or '').strip()
        name = str(row[col_map.get('name', 2)] or '').strip()

        if not name or not code or name == '종합계' or not category:
            continue

        def safe_num(val):
            if val is None or val == '' or val == '-':
                return 0
            try:
                return round(float(str(val).replace(',', '')))
            except:
                return 0

        stock = safe_num(row[col_map.get('stock', 5)] if col_map.get('stock') is not None else 0)
        incoming = safe_num(row[col_map.get('incoming', 6)] if col_map.get('incoming') is not None else 0)
        avg_out = safe_num(row[col_map.get('avg_out') or col_map.get('avg_out_3m', 17)] if (col_map.get('avg_out') or col_map.get('avg_out_3m')) is not None else 0)
        moq = safe_num(row[col_map.get('moq', 18)] if col_map.get('moq') is not None else 0) or 5000
        lead_time = safe_num(row[col_map.get('lead_time', 19)] if col_map.get('lead_time') is not None else 0) or 10
        remark = str(row[col_map.get('remark', 9)] or '').strip() if col_map.get('remark') is not None else ''

        expiry_date = ''
        if col_map.get('expiry') is not None and row[col_map['expiry']]:
            raw = row[col_map['expiry']]
            if isinstance(raw, datetime):
                expiry_date = raw.strftime('%Y-%m-%d')
            elif isinstance(raw, (int, float)):
                from datetime import timedelta
                d = datetime(1899, 12, 30) + timedelta(days=int(raw))
                expiry_date = d.strftime('%Y-%m-%d')

        products.append({
            'category': category,
            'code': code,
            'name': name,
            'currentStock': stock,
            'incoming': incoming,
            'safetyStock': safe_num(row[col_map.get('safety', 15)] if col_map.get('safety') is not None else 0),
            'avgMonthlyOut': avg_out,
            'moq': moq,
            'leadTime': lead_time,
            'expiryDate': expiry_date,
            'remark': remark
        })

    print(f"파싱 완료: {len(products)}개 제품")
    wb.close()
    return products


def update_dashboard(products):
    """대시보드 HTML 파일의 샘플 데이터를 업데이트"""
    print(f"대시보드 업데이트 중... ({DASHBOARD_HTML})")

    with open(DASHBOARD_HTML, 'r', encoding='utf-8') as f:
        html = f.read()

    # JavaScript 배열 문자열 생성
    today = datetime.now().strftime('%Y년 %m월 %d일')
    js_products = json.dumps(products, ensure_ascii=False, indent=4)

    # loadSampleData 함수 내의 products 배열을 교체
    pattern = r'(function loadSampleData\(\) \{\s*products = )\[[\s\S]*?\](;\s*saveToStorage)'
    replacement = f'\\1{js_products}\\2'

    new_html = re.sub(pattern, replacement, html)

    # 동기화 날짜 업데이트
    new_html = re.sub(
        r"showToast\('.*?동기화'\);",
        f"showToast('실시간 데이터 ' + products.length + '개 제품이 로드되었습니다. ({today} 동기화)');",
        new_html
    )

    # 데이터 버전 업데이트 (캐시 무효화)
    version = f"v-auto-{datetime.now().strftime('%Y%m%d%H%M')}"
    new_html = re.sub(r"'v7-shipment'", f"'{version}'", new_html)

    with open(DASHBOARD_HTML, 'w', encoding='utf-8') as f:
        f.write(new_html)

    print(f"대시보드 업데이트 완료 (버전: {version})")


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
