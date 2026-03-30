# Contributing / コントリビューションガイド

Team Calendar プロジェクトへの貢献に興味を持っていただきありがとうございます！

## 🐛 バグ報告

[Issue](https://github.com/freyWylfred/TeamCalendar/issues) から報告してください。以下の情報を含めていただけると助かります：

- OS のバージョン（Windows 10 / 11）
- .NET のバージョン（`dotnet --version` の出力）
- Outlook のバージョン
- 再現手順
- デバッグログの内容（🔍 デバッグログを ON にして取得）

## 💡 機能リクエスト

新しい機能のアイデアがあれば [Issue](https://github.com/freyWylfred/TeamCalendar/issues) で提案してください。

## 🔧 開発環境のセットアップ

### 必要なもの

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- Visual Studio 2022 (17.14+) または VS Code
- Microsoft Outlook（デスクトップ版）

### ビルド

```bash
git clone https://github.com/freyWylfred/TeamCalendar.git
cd TeamCalendar
dotnet build
```

### 実行

```bash
dotnet run --project TeamCalendar
```

## 📝 Pull Request の手順

1. このリポジトリを **Fork** します
2. フィーチャーブランチを作成します
   ```bash
   git checkout -b feature/your-feature-name
   ```
3. 変更を加えてコミットします
   ```bash
   git commit -m "feat: 説明"
   ```
4. フォーク先にプッシュします
   ```bash
   git push origin feature/your-feature-name
   ```
5. **Pull Request** を作成します

### コミットメッセージの規約

[Conventional Commits](https://www.conventionalcommits.org/) に従ってください：

- `feat:` — 新機能
- `fix:` — バグ修正
- `docs:` — ドキュメントの変更
- `refactor:` — リファクタリング
- `style:` — コードスタイルの変更（動作に影響なし）

## 📄 ライセンス

コントリビューションは [MIT License](LICENSE) の下でライセンスされます。
