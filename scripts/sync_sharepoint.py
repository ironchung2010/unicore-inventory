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
    "General/12. SC/F'cst/UNI&CORE Inventory Report_(3월).xlsb")
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
                    # GUID 매칭 숥패 시 최신 Inventory Report 파일 사용
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
            file_path = "General/12. SC/F'cst/UNI&CORE Inventory Report_(3월).xlsb"

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
