# 🚀 Code2PPTX-AI-Analyzer

## 📌 概要（Overview）
このシステムは、対象となるPythonソースコード群を静的解析（AST）およびAI（Gemini API）によって自動で深く読み解き、その解析結果をもとに**プロフェッショナル仕様のPowerPoint（PPTX）コード解説・研修レポート資料を完全自動生成**するPythonツールです。  
手作業によるドキュメント作成やコード解説資料の構成検討にかかる工数を劇的に削減します。

## 🎯 主な目的（Objective）
- **ドキュメント作成の自動化**: コードの読み込みからスライド構成・レイアウト描画までの完全自動パイプライン化
- **ハレーション・捏造の防止**: 温度パラメータ（Temperature=0.0）や厳格なASTフォールバックを組み合わせた、事実に基づく高精度なコード解析
- **研修・レビューの効率化**: 処理フロー、主要機能の解説、主要ライブラリ解説辞書を含んだ高品質な成果物の即時出力

## 🛠️ 使用技術（Tech Stack）
- **言語**: Python
- **AI・検証**: Google Gemini API (`gemini-2.0-flash-lite`), Pydantic（構造化スキーマ検証）, tenacity（自動リトライ制御）
- **主要ライブラリ**: `python-pptx`, `openpyxl`, `pandas`, `ast` (標準ライブラリ) など
- **その他**: バッチ処理スクリプト（`.bat`）によるワンクリック一括実行対応

## ⚙️ システム構成・処理フロー（Workflow）
- **Step 1 (分析・構造化)**: inputs/ 内のPythonコードを読み込み、構文検証（AST）およびGemini API（または厳格なASTフォールバック）を用いてプログラムの目的、処理フロー、主要機能を構造化JSONデータへ変換
- **Step 2 (PPTX自動描画)**: 変換された構造化データをテンプレートに流し込み、表紙、概要・実行フロー、実装解説、用語解説辞書（Appendix）を含むPowerPointレポート（.pptx）を自動描画・保存

## 🚀 使い方（Quick Start）
- **💻 1. 開発者・セットアップ手順**:
- **1.リポジトリのクローン**: 「git clone [https://github.com/mochapaparocku-crypto/Code2PPTX-AI-Analyzer.git](https://github.com/mochapaparocku-crypto/Code2PPTX-AI-Analyzer.git)」
- **2.依存ライブラリのインストール**: 「pip install -r requirements.txt」
- **3.環境変数（.env）の設定（Gemini APIを利用する場合）**: 「GEMINI_API_KEY=your_api_key_here」

- **👤 2. ツールの実行フロー**:
- **1.**: inputs/ フォルダ内に解析したいPythonスクリプト（.py）を配置します[cite: 3]。
- **2.**: 実行用バッチ、またはコマンドを起動します。
- **2-1.**: バッチで実行する場合: run_main_code2pptx.bat をダブルクリック[cite: 3]
- **2-2.**: コマンドで実行する場合:「python system/main_code2pptx.py」
- **3.**: 処理が完了すると、outputs/ フォルダ内に解析用JSONと見栄えの整ったPowerPoint研修レポート（成果報告_コード解説_*.pptx）が出力されます[cite: 5, 6]。

## 📁 フォルダ構成（Repository Structure）
```text
Code2PPTX-AI-Analyzer/
├── inputs/                  # 解析対象のPythonスクリプト群（.py）を格納
├── outputs/                 # 自動生成された中間JSONデータおよびPPTXレポートの出力先
├── templates/               # レポート出力用のPowerPointマスターテンプレート（aspark_org.pptx）
├── system/                  # メインシステムプログラム群
│   ├── main_code2pptx.py    # 全体処理を一元管理する一括実行パイプライン[cite: 6]
│   ├── analysis_python2json.py # コードのAST/AI解析とJSON構造化モジュール[cite: 4]
│   └── generate_json2pptx.py # 構造化データに基づくPPTX自動描画エンジン[cite: 5]
├── run_main_code2pptx.bat   # 文字化け防止・環境チェック付き一括実行バッチ
└── .gitignore               # 機密情報・成果物・キャッシュファイルの除外設定
