import os
import json
import csv
import requests
from datetime import datetime
from zoneinfo import ZoneInfo

# 環境変数から認証情報を取得
ACCESS_TOKEN = os.environ.get("TEAMS_ACCESS_TOKEN", "")
COOKIE_TOKEN = os.environ.get("TEAMS_COOKIE_TOKEN", "")

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

def get_user_info(access_token: str, user_id: str) -> tuple[str, str]:
    """
    creator から displayName と userPrincipalName を取得
    存在しない user_id や認証エラー時には空文字を返却
    """
    url = f"https://teams.microsoft.com/api/mt/apac/beta/users/{user_id}/"
    headers = {
        "accept": "application/json, text/plain, */*",
        "Authorization": f"Bearer {access_token}",
    }
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Failed to fetch user info for {user_id}: {e}")
        return "", ""

    data = response.json()
    user_info = data.get('value', {})
    return user_info.get('displayName', ''), user_info.get('userPrincipalName', '')

def main():
    # JSON 取得
    data = get_customemoji(ACCESS_TOKEN)
    emoticons = data.get('categories', [])[0].get('emoticons', [])

    # 事前に全creator UUIDをユニークに抽出
    creator_ids = set()
    for e in emoticons:
        raw = e.get('creator', '')
        if raw and ':' in raw:
            creator_ids.add(raw.split(':')[-1])

    # 全ユーザー情報を事前取得してキャッシュ
    user_cache: dict[str, tuple[str, str]] = {}
    for uid in creator_ids:
        display, upn = get_user_info(ACCESS_TOKEN, uid)
        user_cache[uid] = (display, upn)

    # CSV 用リスト
    rows = []
    for e in emoticons:
        doc_id = e.get('documentId', '')
        file_name = f"{doc_id}.png"

        # ダウンロード
        png_url = f"https://jp-prod.asyncgw.teams.microsoft.com/v1/objects/{doc_id}/views/imgt2_anim?v=1"
        download_png(png_url, file_name, COOKIE_TOKEN)

        # 作成日時変換
        ts_ms = e.get('createdOn')
        if isinstance(ts_ms, (int, float)):
            dt = datetime.fromtimestamp(ts_ms / 1000, tz=ZoneInfo('Asia/Tokyo'))
            created_on = dt.strftime('%Y-%m-%d %H:%M:%S %Z')
        else:
            created_on = ''

        # creator UUID
        raw_creator = e.get('creator', '')
        creator_id = raw_creator.split(':')[-1] if raw_creator and ':' in raw_creator else ''
        disp_name, upn = user_cache.get(creator_id, ('', ''))

        rows.append({
            'id': e.get('id', ''),
            'documentId': doc_id,
            'description': e.get('description', ''),
            'fileName': file_name,
            'creatorRaw': raw_creator,
            'creator': creator_id,
            'creatorDisplayName': disp_name,
            'creatorUPN': upn,
            'createdOn': created_on,
            'isDeleted': e.get('isDeleted', False)
        })

    # CSV出力
    fieldnames = [
        'id','documentId','description','fileName',
        'creatorRaw','creator','creatorDisplayName','creatorUPN',
        'createdOn','isDeleted'
    ]
    with open('customemoji.csv', 'w', newline='', encoding='utf-8') as csvf:
        writer = csv.DictWriter(csvf, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    # JSON 保存
    with open('customemoji_with_files.json', 'w', encoding='utf-8') as jf:
        json.dump(data, jf, ensure_ascii=False, indent=2)

    print("Downloaded images and exported CSV: customemoji.csv")

if __name__ == '__main__':
    main()
