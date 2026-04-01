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

import msal
import requests

# ============ 설정 ============
TENANT_ID = os.environ.get("AZURE_TENANT_ID", "")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")
SITE_NAME = "msteams_b2be8e"
SHAREPOINT_HOST = "uniandcore.sharepoint.com"
FILE_PATH = os.environ.get("SHAREPOINT_FILE_PATH",
    "/General/12. SC/F'cst/UNI&CORE Inventory Report_(3월).xlsb")
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
    site_id = resp.json()["id"]

    # 드라이브 (문서 라이브러리) 조회
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
        print("ERROR: 문서 라이브러리를 찾을 수 없습니다.")
        sys.exit(1)

    # 파일 다운로드
    file_path = FILE_PATH.lstrip("/")
    download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_path}:/content"
    print(f"파일 다운로드 중... ({file_path})")
    resp = requests.get(download_url, headers=headers)

    if resp.status_code != 200:
        print(f"ERROR: 파일 다운로드 실패 ({resp.status_code}): {resp.text}")
        sys.exit(1)

    local_path = "/tmp/inventory_report.xlsb"
    with open(local_path, "wb") as f:
        f.write(resp.content)

    print(f"다운로드 완료: {local_path} ({len(resp.content)} bytes)")
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

