# Changelog

このプロジェクトのすべての注目すべき変更はこのファイルに記録されます。  
フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.0.0/) に準拠しています。

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
