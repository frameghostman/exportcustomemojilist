# Extracting Custom Emoji From Teams!

This script is for downloading a custom emoji list from Microsoft Teams. <br> 

このスクリプトは、Microsoft Teamsからカスタム絵文字リストをダウンロードするためのものです。<br>
creatorに含まれるuuidからdisplayname,UPN表示まで行っています。<br>

## How to use:使用方法

1. Webブラウザ版Teamsを起動し適当なカスタム絵文字を追加すると https://teams.microsoft.com/api/csa/apac/api/v1/customemoji/metadata にアクセスしているのでその時のトークン(のちにTEAMS_ACCESS_TOKENにセット)を得ます。デベロッパーツール等で確認できる authorization フィールドのbearer以下の部分になります。以下の画像部分です。<br>
When you start the web browser version of Teams and add an appropriate custom emoji, you can notice accessing https://teams.microsoft.com/api/csa/apac/api/v1/customemoji/metadata with a token (later set to TEAMS_ ACCESS_TOKEN). This is the part below the bearer of the authorization field, which can be checked in the developer tools. See the image below for reference.

   ![トークンの取得方法](1.png)

2. 次に "https://jp-prod.asyncgw.teams.microsoft.com/v1/objects/0-ejp" が含まれるあて先に通信しているものを見つけ、どれでもよいのでカスタム絵文字が保存されているアドレスを得ます。その後、そのアドレスにアクセスします。以下の画像部分です。<br>
Next, find a destination that contains “https://jp-prod.asyncgw.teams.microsoft.com/v1/objects/0-ejp” and obtain the address where the custom emoji is stored, whichever one you prefer. Then access that address. See the image below for reference.

   ![カスタム絵文字が保存されているアドレス](2.png)

3. カスタム絵文字が保存されているアドレスにアクセスしたら、そのアドレスへのアクセスではなく favicon.ico へのアクセスからCokkieを得ます。Cokkie フィールドの全量になります。以下の画像部分です。<br>
When you access the address where the custom emoji is stored, you need to get the Cokkie from accessing favicon.ico. You need to get the entire part of Cokkie field. See the image below for reference.

   ![Cokkieの取得方法](3.png)

4. 得たトークンとCokkieを環境変数を設定します：<br>
Set Token and Cokkie to environment variables like below.
   ```
   # Windows
   set TEAMS_ACCESS_TOKEN=your_access_token_here
   set TEAMS_COOKIE_TOKEN=your_cookie_token_here
   
   # Linux/Mac
   export TEAMS_ACCESS_TOKEN=your_access_token_here
   export TEAMS_COOKIE_TOKEN=your_cookie_token_here
   ```

5. スクリプトを実行します：<br>
Runs the script.
   ```
   python customemoji.py
   ```

6. 必要に応じて、customemoji.csv を Excel で開いたのち付属のVBAを実行すると画像一覧も読み込めます。<br>
You can merge from customemoji.csv and image set to Excel file with Included VBA.

## Disclaimer:免責事項

このスクリプトは、個人的な使用のためのものです。<br>
This script is for personal use only.