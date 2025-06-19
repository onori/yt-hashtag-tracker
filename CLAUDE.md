# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

**注意**: このプロジェクトは日本語のプロジェクトです。コード内のコメント、ログメッセージ、変数名、および関連する説明は日本語で記述してください。

## プロジェクト概要

特定のハッシュタグ（#安野たかひろ、#チームみらい）が付いたYouTube動画を自動的に追跡し、データをGoogle スプレッドシートに記録するGoogle Apps Scriptプロジェクトです。対象ハッシュタグの動画を検索し、動画のメタデータと統計情報を抽出して、構造化された形式でスプレッドシートに保存します。

## 開発コマンド

### 基本コマンド
- `npm install` - 依存関係をインストール
- `npm run push` - Google Apps Scriptにコードをアップロード
- `npm run pull` - Google Apps Scriptからコードをダウンロード
- `npm run deploy` - コードをプッシュしてデプロイ
- `npm run open` - Google Apps Scriptエディタを開く
- `npm run watch` - TypeScriptファイルの変更を監視

### Google Apps Script セットアップ
- `clasp login` - Google Apps Scriptで認証
- `clasp create --type sheets --title "YouTubeハッシュタグトラッカー" --rootDir ./src` - 新しいGASプロジェクトを作成

### コード品質
- コードフォーマットとリンティングはBiomeで処理（`biome.json`に設定）
- タブインデントとダブルクォートを使用
- コード品質チェックは直接Biomeコマンドを実行

## アーキテクチャ

### コア構造
- **メインエントリーポイント**: `src/index.ts` - 全ての主要機能を含む
- **型定義**: `src/types/youtube.ts` - YouTube APIレスポンスの型
- **設定**: `src/appsscript.json` - GASプロジェクト設定とAPIスコープ

### 主要関数
- `main()` - 動画検索とデータ収集を統括するメイン実行関数
- `dailyUpdate()` - 積み上げ追跡のための日次データ収集
- `fetchYouTubeVideoData()` - YouTube API統合とデータ取得のコア機能
- `removeDuplicateVideos()` - 最新の動画データを保持する重複削除ロジック
- `updateDailyStats()` - 日次統計の集計
- `updateSubscriberHistory()` - チャンネル登録者数の時系列追跡

### データフロー
1. 対象ハッシュタグでYouTube動画を検索（過去365日間）
2. YouTube Data API v3を使用して詳細な動画メタデータを取得
3. チャンネル情報と登録者数を取得
4. コンテンツ分析に基づいて動画を「ショート」または「通常」に分類
5. 重複削除を行いながら構造化データをGoogle スプレッドシートに保存
6. 日次統計と登録者数履歴の追跡を生成

### Google スプレッドシート統合
- 複数シートの作成と管理：メインデータ、日次統計、登録者数履歴
- 自動ヘッダー設定とフォーマット
- 動画IDに基づく重複削除（最新データを保持）
- ヘッダー行の固定と列幅の自動調整

### API依存関係
- YouTube Data API v3：動画検索とメタデータ取得
- Google Sheets API：データ保存
- YouTube読み取り専用とSheetsアクセスのOAuthスコープが必要

### 設定項目
- `src/index.ts`の`TARGET_HASHTAGS`配列：追跡対象ハッシュタグ
- スクリプトプロパティの`SPREADSHEET_ID`：対象スプレッドシート
- `appsscript.json`でタイムゾーンをAsia/Tokyoに設定

## TypeScript設定

- ESモジュールを使用（`package.json`の`"type": "module"`）
- `@types/google-apps-script`でGoogle Apps Scriptの型を含む
- 対象スプレッドシート列：取得日時, ハッシュタグ, 動画ID, 動画カテゴリ, 動画タイトル, 動画URL, チャンネル名, チャンネル登録者数, 動画公開日, 動画の説明, 視聴回数, いいね数, コメント数