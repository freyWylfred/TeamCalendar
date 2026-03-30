# Release Notes v1.3.0

## 🎯 単一 exe 配布
- **全 DLL を exe に結合** — `TeamCalendar.exe` 1つ + `config.ini` のみで動作
- **.NET ランタイム同梱** — Windows 10/11 で追加インストール不要
- **圧縮有効** — 約 46 MB（ランタイム込み）

## 📦 配布内容
| ファイル | 説明 |
|---|---|
| `TeamCalendar.exe` | アプリ本体（DLL + ランタイム結合済み） |
| `config.ini` | 勤務時間設定（初回起動時に自動生成） |

## 使い方
1. zip を任意のフォルダに展開
2. `TeamCalendar.exe` を実行
3. 必要に応じて `config.ini` で勤務時間を編集
