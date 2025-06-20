# 【完全版】GoogleカレンダーLINEbot作成ガイド - つまずきポイント全解説

## はじめに

GoogleカレンダーをLINEに自動送信するbotを作ってみました。簡単そうに見えて意外とつまずきポイントが多かったので、実際に遭遇したエラーと解決方法を含めて詳しく解説します。

**完成すると何ができるか：**
- 毎朝8時に今日の予定をLINEに自動送信
- LINEで「今日の予定」と送ると即座に予定を返信
- スプレッドシートから簡単に手動送信も可能

**所要時間：** 約30分（エラー対応含む）

> 📊 この記事の開発過程と技術的な背景については、[こちらの資料](https://skywork.ai/share/v2/ppt/1935948740639268864?pid=1935946615564980224&sid=gen_ppt-p0ojMqriy&t=gen_ppt)も合わせてご参照ください。

---

## 📋 事前準備

### 必要なもの
- Googleアカウント
- LINEアカウント（Developersアカウント作成用）
- スマートフォン（User ID取得用）

### 参考情報
今回作成するbotの構成：
```
Googleカレンダー → Google Apps Script → LINE Messaging API → LINEアプリ
```

---

## Step 1: LINE Developers設定

### 1-1. アカウント作成とプロバイダー設定

1. [LINE Developers](https://developers.line.biz/) にアクセス
2. LINEアカウントでログイン
3. 初回の場合は開発者登録
4. 「Create」→「Create a new provider」
5. プロバイダー名を入力（例：「MyBotProvider」）

### 1-2. Messaging APIチャンネル作成

1. 作成したプロバイダーをクリック
2. 「Create a new channel」→「Messaging API」を選択
3. 必要情報を入力：
   - **Channel name**: カレンダーbot（任意）
   - **Channel description**: Googleカレンダー連携bot
   - **Category**: Tools
   - **Subcategory**: その他
4. 利用規約に同意して「Create」

### 1-3. 重要情報を取得

チャンネル作成後、以下の情報をメモ：

1. **Channel Access Token（長期）**
   - 「Messaging API設定」タブ
   - 「Channel access token」の「Issue」をクリック
   - 生成されたトークンをコピー

2. **Channel ID**
   - 「基本設定」タブに表示

3. **Channel Secret**
   - 「基本設定」タブに表示

**⚠️ 注意：** これらの情報は後で使用するので必ずメモしてください。

---

## Step 2: Google Apps Script設定

### 2-1. GASプロジェクト作成

1. [Google Apps Script](https://script.google.com/) にアクセス
2. 「新しいプロジェクト」をクリック
3. プロジェクト名を「カレンダーLINEbot」に変更

### 2-2. マニフェストファイル設定

**重要：** 最初にこれを設定しないと後でエラーになります。

1. 「ファイル」→「プロジェクトのプロパティ」→「マニフェストファイルをエディタで表示する」にチェック
2. `appsscript.json` ファイルが表示される
3. 内容を以下に置き換え：

```json
{
  "timeZone": "Asia/Tokyo",
  "dependencies": {
    "enabledAdvancedServices": []
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "ANYONE"
  },
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/script.external_request"
  ]
}
```

### 2-3. メインスクリプト設定

1. `Code.gs` の内容を全て削除
2. [提供されたCode.gsのコード]をコピー&ペースト
3. 「ファイル」→「新規作成」→「スクリプトファイル」で `Utils.gs` を作成
4. [提供されたUtils.gsのコード]をペースト

### 2-4. スプレッドシート作成

1. 「リソース」→「ライブラリ」→「このスクリプトに関連付けられたスプレッドシートを作成」
2. スプレッドシートが開いたら名前を「カレンダーLINEbot設定」に変更

---

## Step 3: 初期設定と権限許可

### 3-1. 権限許可

**⚠️ つまずきポイント：** 権限許可は段階的に行う必要があります。

1. GASエディタで `onOpen` 関数を選択
2. 「実行」ボタンをクリック
3. 権限許可画面が表示される
4. 「詳細」→「〜(安全ではないページ)に移動」をクリック
5. 権限を確認して「許可」をクリック

**権限の内容：**
- Googleスプレッドシートへのアクセス
- Googleカレンダーの読み取り
- 外部サービス（LINE API）へのリクエスト

### 3-2. スプレッドシート初期化

1. 作成したスプレッドシートを開く
2. メニューバーに「📅 カレンダーLINEbot」が表示されることを確認
3. 「📝 設定シートを初期化」をクリック
4. 「カレンダー設定」シートが作成される

### 3-3. 設定値入力（一部）

まず以下の項目を入力：

| セル | 項目 | 値 | 備考 |
|------|------|-----|------|
| B1 | GoogleカレンダーID | （後で設定） | |
| B2 | LINE Channel Access Token | Step 1で取得したトークン | |
| B3 | LINE User ID | （後で設定） | ⚠️ LINE IDではない |
| B4 | Webhook Verify Token | （空白でOK） | |

**カレンダーID取得方法：**
1. [Googleカレンダー](https://calendar.google.com/) を開く
2. 左サイドバーで使用したいカレンダーの「⋮」→「設定と共有」
3. 「カレンダーの統合」セクションで「カレンダーID」をコピー
4. B1セルに入力

---

## Step 4: デプロイとWebhook設定

### 4-1. Webアプリとしてデプロイ

1. GASエディタで「デプロイ」→「新しいデプロイ」
2. 設定：
   - **種類**: ウェブアプリ
   - **実行者**: 自分
   - **アクセスできるユーザー**: 全員
3. 「デプロイ」をクリック
4. **ウェブアプリのURL**をコピー

**⚠️ 重要：** このURLは後で使用するので必ずメモしてください。

### 4-2. LINE Webhook設定

1. LINE Developers コンソールに戻る
2. 「Messaging API設定」タブを開く
3. **Webhook URL**に Step 4-1 でコピーしたURLを入力
4. **Webhookの利用**: オン
5. **応答メッセージ**: オフ（重要！）

### 4-3. Webhook検証

「検証」ボタンをクリック

**🚨 エラーが発生する場合（よくあります）：**

#### エラー1: 302 Found
```
ボットサーバーから200以外のHTTPステータスコードが返されました。（302 Found）
```

**原因：** GASのレスポンス処理に問題
**解決方法：** 
1. GASエディタで `doPost` 関数と `doGet` 関数が正しく実装されているか確認
2. 再デプロイが必要な場合があります

#### エラー2: 検証タイムアウト
**原因：** デプロイ設定の問題
**解決方法：**
1. 「デプロイを管理」→「編集」
2. 「アクセスできるユーザー」が「全員」になっているか確認

---

## Step 5: User ID取得

**⚠️ 最大のつまずきポイント：** LINE IDとUser IDの違い

### 5-1. LINE IDとUser IDの違い

- **LINE ID**: `@306sgfhi` (人間が覚えやすいID)
- **User ID**: `U1234567890abcdef...` (API用の内部ID、32文字程度)

**API通信にはUser IDが必要です！**

### 5-2. User ID取得方法

1. 作成したLINE botを友達追加（QRコードまたはLINE ID）
2. botに「**userid**」または「**test**」と送信
3. botから返信されるUser IDをコピー
4. スプレッドシートのB3セルに入力

**🚨 よくあるエラー：**
```
https://api.line.me のリクエストに失敗しました（エラー: 400）
The property, 'to', in the request body is invalid
```

**原因：** LINE IDをUser IDの欄に入力している
**解決方法：** 正しいUser IDを取得して入力し直す

---

## Step 6: 最終テストとトラブルシューティング

### 6-1. 設定確認

1. スプレッドシートメニュー「🔍 設定状況を確認」をクリック
2. すべての項目が「✅」になることを確認

### 6-2. 手動送信テスト

1. 「📤 今日のイベントを手動送信」をクリック
2. LINEに今日の予定が送信されることを確認

### 6-3. 双方向通信テスト

LINE botに以下を送信してテスト：
- `今日の予定`: 今日の予定を表示
- `明日の予定`: 明日の予定を表示
- `ヘルプ`: 使い方を表示

---

## よくあるエラーと解決方法

### 1. カレンダー関連エラー

```
指定されたカレンダーIDが見つかりません
```

**解決方法：**
- カレンダーIDが正しいか再確認
- プライベートカレンダーの場合は共有設定を確認
- デフォルトカレンダーの場合はGmailアドレスを使用

### 2. LINE API関連エラー

```
LINE API Error: 401 - Unauthorized
```

**解決方法：**
- Channel Access Tokenが正しいか確認
- トークンの有効期限を確認

```
LINE API Error: 400 - Bad Request
```

**解決方法：**
- User IDが正しいか確認（LINE IDではなくUser ID）
- メッセージ形式に問題がないか確認

### 3. Webhook関連エラー

```
Webhook接続エラー
```

**解決方法：**
1. GASが正しくデプロイされているか確認
2. Webhook URLが正確にコピーされているか確認
3. 「アクセスできるユーザー」が「全員」になっているか確認

### 4. 権限関連エラー

```
Exception: You do not have permission to call ...
```

**解決方法：**
1. GASの権限許可を再実行
2. ブラウザのポップアップブロック設定を確認
3. `appsscript.json` の `oauthScopes` が正しいか確認

---

## 応用設定

### 自動送信設定

毎朝8時に自動送信する場合：
1. スプレッドシートメニュー「⏰ 朝8時の自動送信を設定」をクリック

### カスタマイズ例

1. **送信時間の変更**
   - `Utils.gs` の `setupMorningTrigger` 関数を編集

2. **メッセージ形式の変更**
   - `Code.gs` の `createLineMessage` 函数を編集

3. **追加コマンドの実装**
   - `handleTextMessage` 関数に新しい条件分岐を追加

---

## デバッグ方法

### ログの確認方法

1. GASエディタで「表示」→「ログ」
2. エラーの詳細を確認
3. 「🔍 設定状況を確認」メニューで設定の妥当性をチェック

### よく使うデバッグ関数

```javascript
// カレンダー接続テスト
testGetTodayEvents()

// 設定確認
showSettingsStatus()

// トリガー確認
showCurrentTriggers()
```

---

## セキュリティ考慮事項

### 機密情報の管理

1. **アクセストークン**: スプレッドシートに保存されるため、適切な共有設定を行う
2. **Webhook検証**: セキュリティを高めたい場合は Verify Token を設定
3. **権限管理**: 必要最小限の権限のみを許可

### 推奨設定

- スプレッドシートの共有設定を「リンクを知っている全員が編集可」から「特定のユーザーのみ」に変更
- 定期的にアクセストークンを更新
- ログの定期的な確認

---

## まとめ

GoogleカレンダーLINEbotの作成は、以下のポイントを押さえれば比較的スムーズに進められます：

1. **マニフェストファイルの正しい設定** - 最初が肝心
2. **LINE IDとUser IDの区別** - 最大のつまずきポイント
3. **権限許可の段階的実行** - 焦らず一つずつ
4. **デバッグログの活用** - エラー時は必ずログを確認

完成すると、毎日の予定管理が格段に楽になります。ぜひチャレンジしてみてください！

---

## 参考リンク

- [Google Apps Script](https://script.google.com/)
- [LINE Developers](https://developers.line.biz/)
- [LINE Messaging API リファレンス](https://developers.line.biz/ja/reference/messaging-api/)
- [Google Calendar API](https://developers.google.com/calendar/api)
- [今回の開発で使用した資料・プレゼンテーション](https://skywork.ai/share/v2/ppt/1935948740639268864?pid=1935946615564980224&sid=gen_ppt-p0ojMqriy&t=gen_ppt)

## 更新履歴

- 2024/06/20: 初版作成
- 実際のつまずきポイントとエラー解決方法を反映 