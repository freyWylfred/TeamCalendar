# 📅 Team Calendar

> **[🇬🇧 English README](../README.md)**

Outlook の予定表から会議情報を取得し、承認状況を一覧表示・Excel 出力できる Windows デスクトップアプリケーションです。  
自分だけでなく、他のチームメンバーの共有カレンダーもまとめて取得できます。

![.NET 10](https://img.shields.io/badge/.NET-10.0-purple)
![Windows Forms](https://img.shields.io/badge/UI-Windows%20Forms-blue)
![License: MIT](https://img.shields.io/badge/License-MIT-green)

---

## ✨ 主な機能

| 機能 | 説明 |
|------|------|
| **予定表取得** | Outlook COM 経由で予定表を読み取り（繰り返し予定にも対応） |
| **複数ユーザー対応** | 自分＋他ユーザーの共有カレンダーをまとめて取得 |
| **ステータス判定** | 承認 / 任意(仮) / 辞退 / 主催者 / 未応答 を自動分類 |
| **色分け表示** | ステータスごとに行をカラーリング（承認=緑, 任意=黄, 辞退=赤, 主催=青） |
| **サマリーカード** | 全件・承認・任意・辞退の件数を一目で確認 |
| **Excel 出力** | 承認済み会議のみを `.xlsx` ファイルに出力（ClosedXML 使用） |
| **デバッグログ** | トグルで表示できるリアルタイムログパネル（トラブルシュート用） |

---

## 📸 画面イメージ

```
┌──────────────────────────────────────────────────────────┐
│  📅  Team Calendar                       (アクセントブルー)  │
├──────────────────────────────────────────────────────────┤
│  期間 [2025/03/24] 〜 [2025/03/28]  ▶予定を取得  📊Excel出力 │
│  👥 対象ユーザー [user1@example.com; user2@...]  ☑自分を含める │
├──────────────────────────────────────────────────────────┤
│  📋 全予定  │  ✅ 承認/主催  │  ⏳ 任意(仮)  │  ❌ 辞退    │
│     12     │      8       │      3       │     1      │
├──────────────────────────────────────────────────────────┤
│  ユーザー │ 件名     │ 開始日時  │ 終了日時  │ ステータス   │
│  自分     │ 定例会議  │ 03/24 10:00│ 03/24 11:00│ 承認       │
│  user1@.. │ 1on1     │ 03/24 14:00│ 03/24 14:30│ 任意       │
└──────────────────────────────────────────────────────────┘
```

---

## 🔧 必要な環境

- **OS**: Windows 10 / 11
- **ランタイム**: [.NET 10 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/10.0)
- **Outlook**: Microsoft Outlook（デスクトップ版）がインストール済みであること
- **共有カレンダー**: 他ユーザーの予定を取得する場合は、相手側で予定表の共有権限が設定されていること

---

## 🚀 セットアップ

### ビルドして実行

```bash
git clone https://github.com/freyWylfred/TeamCalendar.git
cd TeamCalendar
dotnet run --project TeamCalendar
```

### リリースビルド

```bash
dotnet publish TeamCalendar -c Release -o ./publish
```

`./publish` フォルダ内の `TeamCalendar.exe` を実行してください。

---

## 📖 使い方

### 1. 予定を取得する

1. **期間** を開始日〜終了日で指定（デフォルトは今週の月〜金）
2. 他のユーザーの予定も取得する場合は **「👥 対象ユーザー」** にメールアドレスをセミコロン区切りで入力
3. **「▶ 予定を取得」** ボタンをクリック

### 2. Excel に出力する

1. 予定を取得後、 **「📊 Excel出力 (承認済み)」** ボタンをクリック
2. 保存先を選択すると、承認済み（承認＋主催者）の会議のみが `.xlsx` ファイルに出力されます

### 3. デバッグログ

- **「🔍 デバッグログ」** チェックボックスをONにすると、画面下部にリアルタイムログが表示されます
- Outlook との通信状況やエラー詳細の確認に使用できます

---

## 🏗 技術スタック

| 技術 | 用途 |
|------|------|
| **.NET 10** (Windows Forms) | UI フレームワーク |
| **Outlook COM Interop** (`dynamic`) | Outlook 予定表へのアクセス |
| **ClosedXML** | Excel (.xlsx) ファイル出力 |

---

## 📁 プロジェクト構成

```
TeamCalendar/
├── TeamCalendar.slnx           # ソリューションファイル
├── .gitignore
├── LICENSE
├── README.md                   # 英語版
├── docs/
│   └── README.ja.md            # 日本語版（このファイル）
└── TeamCalendar/
    ├── TeamCalendar.csproj     # プロジェクト定義 (.NET 10)
    ├── Program.cs              # エントリポイント
    ├── Form1.cs                # メインフォーム（ロジック）
    ├── Form1.Designer.cs       # メインフォーム（UI 定義）
    └── Form1.resx              # リソースファイル
```

---

## ⚠️ 注意事項

- Outlook の **デスクトップ版** が必要です（Web 版 Outlook / new Outlook には非対応）
- 他ユーザーの予定表を取得するには、Exchange / Microsoft 365 環境で **予定表の共有アクセス権** が必要です
- 他ユーザーの予定で表示される「ステータス」は、**そのユーザー自身の応答状態** です

---

## 🤝 コントリビューション

Issue や Pull Request は歓迎します。

1. このリポジトリを Fork
2. フィーチャーブランチを作成 (`git checkout -b feature/amazing-feature`)
3. 変更をコミット (`git commit -m 'Add amazing feature'`)
4. ブランチを Push (`git push origin feature/amazing-feature`)
5. Pull Request を作成

---

## 📄 ライセンス

[MIT License](../LICENSE) の下で公開されています。
