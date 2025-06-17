# カスタム絵文字リストエクスポート

このスクリプトは、Microsoft Teamsからカスタム絵文字リストをダウンロードするためのものです。

## 使用方法

1. 環境変数を設定します：
   ```
   # Windows
   set TEAMS_ACCESS_TOKEN=your_access_token_here
   set TEAMS_COOKIE_TOKEN=your_cookie_token_here
   
   # Linux/Mac
   export TEAMS_ACCESS_TOKEN=your_access_token_here
   export TEAMS_COOKIE_TOKEN=your_cookie_token_here
   ```

2. スクリプトを実行します：
   ```
   python customemoji.py
   ```

3. 必要に応じて、customemoji.csv を Excel で開いたのち付属のVBAを実行すると画像一覧も読み込めます。

## 免責事項

このスクリプトは、個人的な使用のためのものです。