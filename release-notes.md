# Release Notes v1.5.0

## 🔗 Exchange Server EWS 認証対応

### Exchange Server へのダイレクト接続
メニュー **[ファイル] > [Exchange接続設定]** から Exchange Server に直接接続できるようになりました。  
Outlook のインストールや共有設定に依存せず、EWS (Exchange Web Services) 経由で予定表を取得します。

**対応認証方式:**
- NTLM 認証（Windows ドメイン環境）
- Basic 認証（ユーザー名 + パスワード）
- ドメイン指定
- 自己署名証明書の SSL 検証スキップ（テスト環境向け）

### ExchangeLoginDialog — 接続情報ダイアログ
- サーバー URL / メールアドレス / パスワード / ドメイン / SSL スキップを入力
- **⚡ 接続テスト**ボタンでダイアログ内から即時疎通確認（✅ / ❌ 表示）
- **🔗 接続**で EWS モードへ切り替え、**🔌 切断**で Outlook COM モードへ戻す

### デュアルモードアーキテクチャ
| モード | 接続方法 | 用途 |
|---|---|---|
| 📧 Outlook COM | ローカル Outlook プロファイル | Exchange 設定不要な環境 |
| 🔗 Exchange EWS | サーバー直接接続 | Outlook 未インストール / 他ユーザー予定表取得 |

ツールバーに現在の接続モードをリアルタイム表示。

### EWS CalendarView
- 定期アイテムを自動展開（繰り返し会議も個別スロットとして表示）
- 複数ユーザーの取得失敗を一括収集し、詳細付きで MessageBox 表示

## 📦 配布内容
| ファイル | 説明 |
|---|---|
| `TeamCalendar.exe` | アプリ本体（DLL + .NET ランタイム結合済み、単一ファイル） |
| `config.ini` | 勤務時間設定（初回起動時に自動生成） |
| `app.ico` | アプリアイコン |

## 使い方
1. zip を任意のフォルダに展開
2. `TeamCalendar.exe` を実行
3. Exchange 接続する場合: メニュー [ファイル] > [Exchange接続設定] で接続情報を入力
4. 必要に応じて `config.ini` で勤務時間を編集
