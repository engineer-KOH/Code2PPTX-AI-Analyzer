"""
===============================================================================
[更新履歴]
- 2026-10-07 v1.3.0 : ハレーション・捏造防止の抜本的改修
  - コメント文の保持（ビジネス目的の解読精度向上）
  - ハレーション防止（ネガティブプロンプトの強化）
  - ASTフォールバックの厳密化（実在する要素のみ抽出）
===============================================================================
"""

import os
import sys
import json
import re
import ast
import time
import logging
from pathlib import Path
from dotenv import load_dotenv
from pydantic import BaseModel, Field

# tenacity による堅牢なリトライ処理
from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    retry_if_exception_type,
    before_sleep_log
)

# ---------------------------------------------------------
# ロギング設定
# ---------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)

# ---------------------------------------------------------
# Gemini API ライブラリの読み込み
# ---------------------------------------------------------
try:
    from google import genai
    from google.genai import types
    from google.genai.errors import APIError
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False
    APIError = Exception

# ---------------------------------------------------------
# 環境変数・パス設定 (system/ からの実行に完全対応)
# ---------------------------------------------------------
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.0-flash-lite")

# コンソール文字化け対策
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# ディレクトリ定義
CURRENT_FILE = Path(__file__).resolve()
PROJECT_ROOT = CURRENT_FILE.parent.parent if CURRENT_FILE.parent.name == 'system' else CURRENT_FILE.parent
INPUT_DIR = PROJECT_ROOT / 'inputs'
OUTPUT_DIR = PROJECT_ROOT / 'outputs'

INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------
# Pydantic による構造化データの型定義
# ---------------------------------------------------------
class BusinessSummary(BaseModel):
    purpose_and_value: str = Field(description="このプログラムの処理内容について、どのような業務上の役割を果たすか（コードの内容のみに基づき記述）")
    target_user: str = Field(description="主な想定利用者（例: 日報作成者、一般社員、システム管理者など）")

class ProcessingStep(BaseModel):
    step_number: int = Field(description="ステップ番号")
    phase_title: str = Field(description="処理フェーズ名")
    input_data: str = Field(description="入力となるデータやファイル")
    process_detail: str = Field(description="実際にコードが行っている具体的な処理内容")
    output_data: str = Field(description="このステップで生成・出力される成果物")

class SectionDetail(BaseModel):
    section_number: int = Field(description="セクション番号")
    title: str = Field(description="【厳守】処理の短縮見出し・機能名（例: マスターファイルの動的特定、カレンダー計算と営業日抽出 など 15文字前後）")
    summary_plain: str = Field(description="【厳守】何をしているかの簡潔な概要（20〜40文字程度）")
    technical_point: str = Field(description="【厳守】実際に使用されている関数名・処理ロジックのポイント")

class LearningPoint(BaseModel):
    topic: str = Field(description="コード内で実際に使われている技術要素・ライブラリ・概念")
    understanding_level: str = Field(description="その技術がどのように活用されているかの具体的な解説")

class DictItem(BaseModel):
    category: str = Field(description="分類（標準ライブラリ、外部ライブラリ、組込み関数など）")
    code_expression: str = Field(description="該当するコード表現や関数名・モジュール名")
    explanation: str = Field(description="このプログラム内での実際の役割や使い方の解説")

class CodeAnalysisResult(BaseModel):
    program_name: str = Field(description="対象プログラム名")
    business_summary: BusinessSummary = Field(description="ビジネス視点での全体概要")
    main_flow: list[ProcessingStep] = Field(description="処理の全体フロー")
    sections: list[SectionDetail] = Field(description="主要機能ごとの詳細解説")
    learning_points: list[LearningPoint] = Field(description="研修・学習理解ポイント")
    dictionary_items: list[DictItem] = Field(description="巻末付録：主要コード・ライブラリ解説辞書")

class BatchCodeAnalysisResponse(BaseModel):
    results: list[CodeAnalysisResult] = Field(description="全解析対象ファイルの解析結果リスト")


# ---------------------------------------------------------
# コードの前処理 (構文検証のみ行い、コメントは保持)
# ---------------------------------------------------------
def validate_and_preprocess_code(file_path: Path) -> str:
    """構文チェック(ast)を行い、無駄な連続空行のみ詰めてコード（コメント含む）を返す"""
    raw_code = file_path.read_text(encoding='utf-8', errors='replace')
    
    try:
        ast.parse(raw_code, filename=file_path.name)
    except SyntaxError as e:
        raise ValueError(f"構文エラー(SyntaxError)が検出されました [{e.filename}:{e.lineno}]: {e.msg}")

    # コメント行は業務の目的が詰まっているため削除せず保持する
    lines = raw_code.splitlines()
    cleaned_lines = []
    
    for line in lines:
        # 空行の連続のみスキップ
        if not line.strip() and cleaned_lines and not cleaned_lines[-1].strip():
            continue
        cleaned_lines.append(line)
        
    cleaned_code = "\n".join(cleaned_lines)
    logging.info(f"前処理完了: {file_path.name} (構文チェック正常・コメント保持)")
    
    return cleaned_code


# ---------------------------------------------------------
# API呼び出し ＆ tenacity 自動リトライ
# ---------------------------------------------------------
@retry(
    retry=retry_if_exception_type((APIError, Exception)),
    stop=stop_after_attempt(5),                          # 最大5回試行
    wait=wait_exponential(multiplier=2, min=4, max=30),  # 指数バックオフ
    before_sleep=before_sleep_log(logging.getLogger(), logging.WARNING),
    reraise=True
)
def _call_gemini_api_batch_with_retry(client: genai.Client, model_name: str, prompt: str) -> str:
    """Gemini API一括呼び出し専用関数"""
    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema=BatchCodeAnalysisResponse,
        temperature=0.0,  # 捏造を排除するためランダム性をゼロにする
        automatic_function_calling=types.AutomaticFunctionCallingConfig(disable=True)
    )
    
    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
        config=config,
    )
    
    if not response.text:
        raise ValueError("Gemini APIからのレスポンスが空でした。")
        
    return response.text


# ---------------------------------------------------------
# 統合解析ルーチン（プロンプト厳格化）
# ---------------------------------------------------------
def analyze_all_python_codes_batch(file_paths: list[Path]) -> list[CodeAnalysisResult]:
    """すべてのPythonコードをまとめたプロンプトを作成し、一括解析する"""
    if not GENAI_AVAILABLE or not GEMINI_API_KEY:
        logging.warning("Gemini APIの設定が未完了です。ローカル厳密AST解析フォールバックを実行します。")
        return [_create_fallback_result(fp) for fp in file_paths]

    cleaned_codes_payload = []
    for fp in file_paths:
        try:
            cleaned = validate_and_preprocess_code(fp)
            cleaned_codes_payload.append(f"=== [ファイル名: {fp.name}] ===\n{cleaned}\n")
        except Exception as e:
            logging.error(f"ファイル前処理失敗 ({fp.name}): {e}")

    if not cleaned_codes_payload:
        return []

    combined_code_text = "\n".join(cleaned_codes_payload)

    client = genai.Client(
        api_key=GEMINI_API_KEY,
        http_options={"api_version": "v1"}
    )

    prompt = f"""
あなたは提示されたコードのみを事実に基づいて誠実に解読するPythonエキスパートです。
以下に含まれるPythonソースコードを解析し、プログラムごとの正確な構造化データを作成してください。

[対象ソースコード群]:
{combined_code_text}

[厳格な絶対ルール（違反厳禁）]:
1. 【捏造・想像の絶対禁止】
   【厳守】コード内に実装されていない機能（例：存在しないWeb API通信、LLMによる考察生成、パレート図作成、tenacityリトライ処理、AST解析など）をの記述を絶対に禁止します。
   【厳守】コード内に書かれている処理、関数、ライブラリ（import文）、変数、コメント文のみを100%事実の根拠としてください。
2. 【プログラム個別の目的の正確な特定】
   【厳守】プログラムのファイル名や処理内容（例：Excel操作、メッセージボックス表示、カレンダー計算など）を正しく読み取り、そのツール本来の目的（例：日報Excelファイルの自動初期化など）を正確に読み取ってください。
3. 【辞書（DictItem）の正確性】
   【厳守】実際に import されている標準モジュールやサードパーティ製モジュール（例：openpyxl など）のみを対象とし、コード内で果たしている役割（関数やメソッドレベル）を具体的に抽出してください。
4.【sections（主要機能）の見出し短縮の徹底】
   【厳守】「■」の行（タイトル・見出し）は、長文を絶対に禁止します。スライドの箇条書きとして一目で認識できるよう、必ず「20文字前後の端的な名詞止め」にしてください。
   【厳守】• 【禁止語句】 見出しの末尾に「〜です」「〜ます」「〜の処理」「〜の機能」を使用することを厳禁とします。
   
[出力形式]:
各ファイルごとに CodeAnalysisResult を作成し、全ファイルをリストで出力してください。
"""

    logging.info(f"全{len(file_paths)}件のファイルをGemini APIへ送信します (モデル: {GEMINI_MODEL}, Temp: 0.0)...")

    try:
        raw_json_text = _call_gemini_api_batch_with_retry(client, GEMINI_MODEL, prompt)
        batch_response = BatchCodeAnalysisResponse.model_validate_json(raw_json_text)
        return batch_response.results
    except Exception as e:
        logging.error(f"[API ERROR] API呼び出しに失敗しました: {e}。ローカル厳密AST解析に切り替えます。")
        return [_create_fallback_result(fp) for fp in file_paths]


def _create_fallback_result(file_path: Path) -> CodeAnalysisResult:
    """API失敗時にコードから実在する要素のみを厳密に抽出するASTフォールバック処理"""
    try:
        source_code = file_path.read_text(encoding="utf-8")
        tree = ast.parse(source_code)
    except Exception as e:
        tree = ast.Module(body=[], type_ignores=[])
        source_code = ""

    # 実際にimportされているモジュールのみ取得
    imports = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                imports.append(alias.name)
        elif isinstance(node, ast.ImportFrom) and node.module:
            imports.append(node.module)
    imports = sorted(list(set(imports)))

    # 定義されている関数・クラス
    functions = [node for node in tree.body if isinstance(node, ast.FunctionDef)]
    classes = [node for node in tree.body if isinstance(node, ast.ClassDef)]

    # モジュール/関数のDocstringやコメントから説明を取得
    module_doc = ast.get_docstring(tree)
    if not module_doc:
        # コメント行から最初の解説らしき行を取得
        comments = [line.strip("# ").strip() for line in source_code.splitlines() if line.strip().startswith("#")]
        module_doc = comments[0] if comments else f"{file_path.name} の自動処理スクリプト"

    # ファイル特性に応じた実態ベースの説明構築
    if "openpyxl" in imports or "excel" in file_path.name.lower():
        purpose = f"{file_path.name} に基づくExcelデータの読み込み・加工・自動更新処理"
        target_user = "業務担当者 / 日報作成者"
    else:
        purpose = f"{file_path.name} スクリプトによる自動処理"
        target_user = "システム管理者 / 開発者"

    # 処理ステップ（実際に定義されている関数から構成）
    main_flow = []
    if functions:
        for idx, fn in enumerate(functions, start=1):
            fn_doc = ast.get_docstring(fn) or "定義された主要ロジックの実行"
            main_flow.append(
                ProcessingStep(
                    step_number=idx,
                    phase_title=f"関数: {fn.name}()",
                    input_data=file_path.name,
                    process_detail=fn_doc.split("\n")[0],
                    output_data="処理結果 / ファイル出力"
                )
            )
    else:
        main_flow.append(
            ProcessingStep(
                step_number=1,
                phase_title="メインスクリプト実行",
                input_data=file_path.name,
                process_detail="トップレベルでの一連のシーケンス処理の実行",
                output_data="実行完了"
            )
        )

    # 実在するライブラリの解説辞書
    dictionary_items = []
    for lib in imports:
        dictionary_items.append(
            DictItem(
                category="使用モジュール",
                code_expression=lib,
                explanation=f"プログラム内で {lib} モジュールの機能を利用"
            )
        )

    return CodeAnalysisResult(
        program_name=file_path.name,
        business_summary=BusinessSummary(
            purpose_and_value=purpose,
            target_user=target_user
        ),
        main_flow=main_flow,
        sections=[
            SectionDetail(
                section_number=1,
                title="プログラム構造と利用ライブラリ",
                summary_plain=f"定義関数: {len(functions)}件 / 定義クラス: {len(classes)}件",
                technical_point=f"使用ライブラリ: {', '.join(imports)}"
            )
        ],
        learning_points=[
            LearningPoint(
                topic="モジュール構成と処理フロー",
                understanding_level=f"{file_path.name} における実際の処理ステップと外部ライブラリの活用"
            )
        ],
        dictionary_items=dictionary_items
    )


# ---------------------------------------------------------
# メイン実行ルーチン
# ---------------------------------------------------------
def main():
    target_files = list(INPUT_DIR.glob("*.py"))

    if not target_files:
        logging.warning("inputs/ フォルダ内に解析対象の .py ファイルが見つかりませんでした。")
        sys.exit(0)

    logging.info(f"解析対象ファイル数: {len(target_files)} 件")
    
    # 統合解析ルーチンの実行
    analysis_results = analyze_all_python_codes_batch(target_files)

    # 結果を個別JSONファイルとして出力フォルダーへ保存
    for idx, (target_file, result) in enumerate(zip(target_files, analysis_results), 1):
        output_json_path = OUTPUT_DIR / f"inputs_{target_file.stem}_analysis.json"
        
        with open(output_json_path, "w", encoding="utf-8") as f:
            f.write(result.model_dump_json(indent=2))

        logging.info(f"[{idx}/{len(target_files)}] 成功: JSON結果を保存しました -> {output_json_path.name}")

if __name__ == "__main__":
    main()