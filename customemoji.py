import os
import json
import csv
import requests
from datetime import datetime
from zoneinfo import ZoneInfo
import msal  # MSAL for auth

# 環境変数から認証情報を取得
CLIENT_ID     = os.environ['CLIENT_ID']
CLIENT_SECRET = os.environ['CLIENT_SECRET']
TENANT_ID     = os.environ['TENANT_ID']
ACCESS_TOKEN  = os.environ.get("TEAMS_ACCESS_TOKEN", "")
COOKIE_TOKEN  = os.environ.get("TEAMS_COOKIE_TOKEN", "")

# --- MSAL (Client Credentials Flow) 設定 ---
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES    = ["https://graph.microsoft.com/.default"]
app = msal.ConfidentialClientApplication(
    client_id=CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

def acquire_graph_token() -> str:
    """MSAL でアクセストークンを取得"""
    result = app.acquire_token_for_client(scopes=SCOPES)  # :contentReference[oaicite:6]{index=6}
    if "access_token" not in result:
        raise RuntimeError(f"Token Error: {result.get('error')} {result.get('error_description')}")
    return result["access_token"]

def get_customemoji(access_token: str):
    """
    カスタム絵文字メタデータを取得
    """
    url = "https://teams.microsoft.com/api/csa/apac/api/v1/customemoji/metadata"
    headers = {
        "accept": "application/json, text/plain, */*",
        "Authorization": f"Bearer {access_token}",
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

def download_png(url: str, output_path: str, cookie_token: str):
    """
    URL から PNG をダウンロードして保存
    """
    headers = {
        "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
    }
    cookies = {"skypetoken_asm": cookie_token}
    resp = requests.get(url, headers=headers, cookies=cookies, stream=True)
    resp.raise_for_status()
    with open(output_path, "wb") as f:
        for chunk in resp.iter_content(8192):
            f.write(chunk)

def get_user_info(graph_token: str, user_id: str) -> tuple[str, str]:
    """
    UUID (Object ID) から displayName と userPrincipalName を取得
    """
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}"
    params = {"$select": "displayName,userPrincipalName"}
    headers = {"Authorization": f"Bearer {graph_token}"}
    r = requests.get(url, headers=headers, params=params)
    r.raise_for_status()
    j = r.json()
    return j.get("displayName", ""), j.get("userPrincipalName", "")  # :contentReference[oaicite:7]{index=7}

def main():
    # JSON 取得
    data = get_customemoji(ACCESS_TOKEN)
    emoticons = data.get('categories', [])[0].get('emoticons', [])
    graph_token = acquire_graph_token()

    # CSV 用リスト
    rows = []

    for e in emoticons:
        doc_id = e.get('documentId')

        # PNG URL 作成
        png_url = f"https://jp-prod.asyncgw.teams.microsoft.com/v1/objects/{doc_id}/views/imgt2_anim?v=1"
        file_name = f"{doc_id}.png"

        # ダウンロード
        download_png(png_url, file_name, COOKIE_TOKEN)

        # JSON レコードにファイル名を追加
        e['fileName'] = file_name

        # createdOn をミリ秒から datetime に変換 (日本時間)
        ts_ms = e.get('createdOn')
        if isinstance(ts_ms, (int, float)):
            dt_jst = datetime.fromtimestamp(ts_ms / 1000, tz=ZoneInfo('Asia/Tokyo'))
            created_on_str = dt_jst.strftime('%Y-%m-%d %H:%M:%S %Z')
        else:
            created_on_str = ''

        # creator フィールドからプレフィックスを除去して UUID 部分を抽出
        raw_creator = e.get('creator', '')

        # 形式例: "8:orgid:0aeaa542-8797-4216-b4cb-e1bb6e72482c"
        if raw_creator and ':' in raw_creator:
            creator_id = raw_creator.split(':')[-1]
        else:
            creator_id = ""

        # UUID 部分が取得できたら Graph API から displayName/UPN を取得
        if creator_id:
            disp_name, upn = get_user_info(graph_token, creator_id)
        else:
            disp_name, upn = "", ""

        # CSV 行作成
        rows.append({
            'id'                  : e.get('id', ''),
            'documentId'          : doc_id,
            'description'         : e.get('description', ''),
            'fileName'            : file_name,
            'creatorRaw'          : raw_creator,
            'creator'             : creator_id,
            'creatorDisplayName'  : disp_name,
            'creatorUPN'          : upn,
            'createdOn'           : created_on_str,
            'isDeleted'           : e.get('isDeleted', '')
        })

    fieldnames = [
        'id','documentId','description','fileName',
        'creatorRaw','creator','creatorDisplayName','creatorUPN',
        'createdOn','isDeleted'
    ]
    with open('customemoji.csv', 'w', newline='', encoding='utf-8') as csvf:
        writer = csv.DictWriter(csvf, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    with open('customemoji_with_files.json','w',encoding='utf-8') as jf:
        json.dump(data, jf, ensure_ascii=False, indent=2)

    print("Downloaded images and exported CSV: customemoji.csv")


if __name__ == '__main__':
    main()