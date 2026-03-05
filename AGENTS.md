# AGENTS.md (FloatVision)

このファイルは、`C:\Users\kasai\source\repos\FloatVision` で作業する AI コーディングエージェント向けのローカル運用ルールです。

## 1. プロジェクト概要

- 種別: Windows 向けフレームレスビューア (Win32 API + C++)
- 主要言語: C++20
- ビルドシステム: Visual Studio C++ (`.vcxproj`)
- 依存関係:
  - `Microsoft.Web.WebView2` (NuGet, `packages.config`)
  - `third_party/md4c` (同梱ソース)

## 2. 変更対象と責務

優先して触るファイル:
- `FloatVision.cpp` (アプリ本体ロジック)
- `FloatVision.rc`, `Resource.h` (リソース/UI 定義)
- `README.md` (ユーザー向け説明)

必要がない限り変更しない:
- `third_party/` 配下 (外部依存)
- `x64/`, `Win32/`, `.vs/` 配下 (生成物/環境依存)
- `packages/` 配下 (NuGet 展開物)

## 3. 作業フロー

1. 変更前に目的と影響範囲を明確化。
2. 小さく変更し、挙動差分を最小化。
3. 変更後に最低 1 つのビルド構成でビルド確認。
4. ユーザーへの報告は以下を必ず含める:
   - 変更ファイル
   - 変更理由
   - 検証結果 (成功/失敗、未実施なら理由)

## 4. ビルドと検証

### 推奨ビルドコマンド (Developer PowerShell)

```powershell
# Debug x64
msbuild .\FloatVision.vcxproj /t:Build /p:Configuration=Debug /p:Platform=x64

# Release x64
msbuild .\FloatVision.vcxproj /t:Build /p:Configuration=Release /p:Platform=x64
```

### NuGet 復元が必要な場合

```powershell
nuget restore .\FloatVision.slnx
```

補足:
- `packages\Microsoft.Web.WebView2...` が無いとビルドエラーになる。
- 実行確認が必要な場合は `x64\Debug\FloatVision.exe` もしくは `x64\Release\FloatVision.exe` を使う。

## 5. コーディング規約 (このリポジトリ用)

- 既存スタイルに合わせる (命名・インデント・改行位置)。
- 不要なリファクタや大規模整形はしない。
- 文字コードや改行コードは既存ファイルに合わせる。
- コメントは「なぜ必要か」を短く書く。自明な処理説明は避ける。

## 6. 安全ルール

- 破壊的操作 (`git reset --hard`, 大量削除など) は、明示的指示がない限り禁止。
- 無関係な差分は戻さない。
- 予期しない既存変更を見つけたら、作業を止めてユーザーに確認する。

## 7. PR/コミット前チェック

- ビルド成功 (少なくとも 1 構成)
- 意図しないファイル変更が無い (`git status`)
- `README.md` との整合性確認 (機能追加/挙動変更時)

## 8. レスポンス方針

- 要点先行で簡潔に報告する。
- バグ修正時は「再現条件」「根本原因」「修正内容」「確認方法」を明記する。
- 未検証項目は推測で断定しない。
