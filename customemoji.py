import requests
import json
import csv
import os
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
        "Content-Type": "application/json",
        "referer": "https://teams.microsoft.com/v2/worker/precompiled-web-worker-7ef81cdf63f16997.js",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Safari/537.36",
        "x-ms-client-type": "web",
        "x-ms-client-version": "1415/25051800219"
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
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
        "Referer": url,
        "Sec-Ch-Ua": '"Google Chrome";v="137", "Chromium";v="137", "Not/A)Brand";v="24"',
        "Sec-Ch-Ua-Mobile": "?0",
        "Sec-Ch-Ua-Platform": '"Windows"',
        "Sec-Fetch-Dest": "image",
        "Sec-Fetch-Mode": "no-cors",
        "Sec-Fetch-Site": "same-origin",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Safari/537.36",
    }
    cookies = {"skypetoken_asm": cookie_token}
    resp = requests.get(url, headers=headers, cookies=cookies, stream=True)
    resp.raise_for_status()
    with open(output_path, "wb") as f:
        for chunk in resp.iter_content(8192):
            f.write(chunk)


def main():
    # JSON 取得
    data = get_customemoji(ACCESS_TOKEN)
    emoticons = data.get('categories', [])[0].get('emoticons', [])

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
        # CSV 行作成
        rows.append({
            'id': e.get('id', ''),
            'documentId': doc_id,
            'description': e.get('description', ''),
            'fileName': file_name,
            'creator': e.get('creator', ''),
            'createdOn': created_on_str,
            'isDeleted': e.get('isDeleted', '')
        })

    # CSV 出力
    with open('customemoji.csv', 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['id', 'documentId', 'description', 'fileName', 'creator', 'createdOn', 'isDeleted']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    # 更新 JSON をファイルに保存（オプション）
    with open('customemoji_with_files.json', 'w', encoding='utf-8') as jf:
        json.dump(data, jf, ensure_ascii=False, indent=2)

    print("Downloaded images and exported CSV: customemoji.csv")

if __name__ == '__main__':
    main()