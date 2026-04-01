# Changelog

このプロジェクトのすべての注目すべき変更はこのファイルに記録されます。  
フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.0.0/) に準拠しています。

## [1.5.0] - 2025-07-26

### Added
- **Exchange Server EWS 認証対応**
  - メニュー [ファイル] > [Exchange接続設定] から Exchange Server へ直接接続可能に
  - EWS (Exchange Web Services) SOAP API を `HttpClient` + `System.Xml.Linq` で実装（追加 NuGet パッケージ不要）
  - NTLM 認証・Basic 認証・ドメイン指定に対応
  - 自己署名証明書の SSL 検証スキップオプション（テスト環境向け）
  - `ExchangeLoginDialog` — 接続情報入力ダイアログ（サーバー URL / メール / パスワード / ドメイン / SSL スキップ）
  - ⚡ 接続テストボタンでダイアログ内から即時接続確認
  - 🔌 切断ボタンで Outlook COM モードへ戻す操作に対応
- **デュアルモードアーキテクチャ**
  - Exchange 接続時 → EWS 経由でダイレクト取得（`LoadCalendarDataViaExchangeAsync`）
  - 未接続時 → 従来の Outlook COM 経由取得（`LoadCalendarData`）
  - ツールバーに接続モード表示ラベル (`📧 Outlook COM` / `🔗 Exchange: xxx@...`) を追加
- **EWS CalendarView** で定期アイテムを自動展開（Outlook COM の `IncludeRecurrences` 相当）
- 他ユーザー取得失敗時を `failedUsers` リストで収集し一括エラー表示（EWS / Outlook COM 共通）

## [1.4.1] - 2025-07-25

### Fixed
- **グラフとチェックボックスの重なり修正**
  - チャート描画の topMargin を拡大し、タイトル/チェックボックスとバーの重なりを解消
  - バー上部のパーセンテージラベルがチャート領域外にはみ出さないようクランプ
  - グラフパネルの高さを 210→240px に拡大
- **サマリーカードとグラフ間の余白を最適化**
  - pnlSummary の高さを 90→78px に縮小、パディングを調整
- **ユーザー取得失敗時のエラー表示を追加**
  - 解決できなかったユーザーの一覧を HRESULT エラーコード・理由と共に MessageBox で表示
  - COMException / Resolve失敗 / その他例外の種別ごとに適切なメッセージを表示

## [1.4.0] - 2025-07-25

### Changed
- **デザイン全面リニューアル** — "Nordic Slate" テーマ
  - ヘッダーを Slate-800→900 グラデーションに変更（高級感のある深いネイビー）
  - アクセントカラーを Indigo-500 (#6366F1) に統一（モダンな紫青）
  - ボタン: Load = Indigo-500, Export = Emerald-600（視認性と統一感）
  - テキスト・ラベルを Slate-500/800 系に統一（目に優しい配色）
  - 背景を Slate-100、カードエリアを Slate-50 に変更（3層の奥行き）
  - サマリーカードに Slate-200 ボーダー + 全高アクセントバーを追加
  - グラフカードにボーダー追加、バー色を Indigo/Emerald に統一
  - タイムライン行色を Emerald-50/Indigo-50/Amber-50/Red-50 のパステルに
  - MenuStripRenderer を Slate-700 ダークテーマに更新

### Fixed
- **例外処理の強化**
  - Form1 コンストラクタ: 設定ファイル読込失敗時に既定値へフォールバック
  - LoadAppIcon: 破損アイコンファイルでのクラッシュ防止
  - ApplyModernTheme: リフレクション呼び出しの保護
  - PaintChart: GDI+ 描画中の例外でクラッシュしないよう try-catch 追加
  - CellDoubleClick / chkIncludeTentative: イベントハンドラの保護
  - WorkScheduleConfig.Load / CreateDefaultIfMissing: ファイル I/O の例外処理
  - MenuStripRenderer: SolidBrush の using 漏れ（GDI+ リソースリーク）を修正

## [1.3.0] - 2025-07-24

### Changed
- 全 DLL を exe に結合した単一ファイル配布に変更（`PublishSingleFile` + `SelfContained`）
- .NET ランタイム同梱により Windows 10/11 で追加インストール不要に
- `config.ini` のパス解決を `Environment.ProcessPath` ベースに改善
- 圧縮有効化によりファイルサイズを最適化

## [1.2.0] - 2025-07-24

### Added
- 会議時間計算ロジックを `MeetingCalculator` クラスに分離（テスト容易化）
- ユニットテストプロジェクト `TeamCalendar.Tests` を追加

### Changed
- `_dayStats` の null 関連処理を `OfType<>()` に置き換え、null 安全性を向上

## [1.1.0] - 2025-07-23

### Added
- 曜日別会議時間/空き時間の積み上げ棒グラフ表示
- タイムラインビュー（勤務時間スロット表示）
- `config.ini` による勤務時間・休憩時間・スロット間隔の設定

## [1.0.0] - 2025-03-30

### Added
- Outlook 予定表から会議情報を取得する機能
- 自分 + 他ユーザーの共有カレンダーをまとめて取得する機能
- 会議の応答ステータス判定（承認 / 任意 / 辞退 / 主催者 / 未応答）
- ステータスごとの色分け表示
- サマリーカード（全件・承認・任意・辞退の件数表示）
- 承認済み会議の Excel (.xlsx) 出力機能
- デバッグログパネル（トグル表示）
- 充実した例外処理（Outlook COM エラー、ファイル I/O エラー等）
- Fluent Design 風モダン UI
