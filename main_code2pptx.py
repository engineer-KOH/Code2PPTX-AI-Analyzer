"""
===============================================================================
[システム一括実行パイプライン] main_code2pptx.py
- Step 1: analysis_python2json.py (Pythonコード -> AI/AST解析 -> JSON出力)
- Step 2: generate_json2pptx.py   (JSONデータ -> PPTX自動描画・生成)
===============================================================================
"""

import os
import sys
import json
import logging
from pathlib import Path

# 新ファイル名からのインポート設定
try:
    from analysis_python2json import analyze_all_python_codes_batch, INPUT_DIR, OUTPUT_DIR
    from generate_json2pptx import render_pptx_from_analysis
except ImportError as e:
    logging.error(f"モジュールのインポートに失敗しました。`analysis_python2json.py` と `generate_json2pptx.py` が同じフォルダにあるか確認してください: {e}")
    sys.exit(1)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


def run_full_pipeline():
    """解析からPPTX生成までを一元管理・一括実行するメイン処理"""
    base_dir = Path(__file__).resolve().parent
    project_root = base_dir.parent if base_dir.name == 'system' else base_dir
    
    inputs_dir = project_root / "inputs"
    outputs_dir = project_root / "outputs"
    template_file = project_root / "templates" / "aspark_org.pptx"

    inputs_dir.mkdir(parents=True, exist_ok=True)
    outputs_dir.mkdir(parents=True, exist_ok=True)

    # 1. inputs/ 内の Python ファイルを検出
    python_files = sorted(list(inputs_dir.glob("*.py")))
    if not python_files:
        logging.warning(f"解析対象の Python ファイルが見つかりません。`{inputs_dir}` フォルダに .py ファイルを配置してください。")
        return

    logging.info(f"==================================================")
    logging.info(f"【パイプライン開始】 検出された対象ファイル数: {len(python_files)} 件")
    logging.info(f"==================================================")

    # 2. Step 1: AI解析 & JSON出力 (analysis_python2json)
    logging.info("▶ [Step 1/2] analysis_python2json による静的・AI解析を実行中...")
    try:
        analysis_results = analyze_all_python_codes_batch(python_files)
    except Exception as e:
        logging.error(f"[Step 1 エラー] 解析処理中にエラーが発生しました: {e}")
        return

    # 3. Step 2: 解析JSONからPPTXレポートを生成 (generate_json2pptx)
    logging.info("▶ [Step 2/2] generate_json2pptx による PowerPoint レポートの描画・生成を実行中...")

    success_count = 0
    for target_file, result_data in zip(python_files, analysis_results):
        try:
            # 中間データ（JSON）の保存処理
            json_filename = f"inputs_{target_file.stem}_analysis.json"
            json_output_path = outputs_dir / json_filename
            
            if hasattr(result_data, "model_dump_json"):
                json_str = result_data.model_dump_json(indent=2)
            elif isinstance(result_data, dict):
                json_str = json.dumps(result_data, ensure_ascii=False, indent=2)
            else:
                json_str = str(result_data)

            json_output_path.write_text(json_str, encoding="utf-8")
            logging.info(f"  ├─ [JSON保存完了]: {json_filename}")

            # PowerPoint レポート生成 (標準化統合関数を呼び出し)
            output_pptx_filename = f"成果報告_コード解説_{target_file.name}.pptx"
            output_pptx_path = outputs_dir / output_pptx_filename

            # 描画エンジンに渡し、PPTX生成を実行
            render_pptx_from_analysis(result_data, template_file, output_pptx_path)

            logging.info(f"  └─ [PPTX生成完了]: {output_pptx_filename}")
            success_count += 1

        except Exception as e:
            logging.error(f"  └─ [エラー] {target_file.name} のレポート生成に失敗しました: {e}", exc_info=True)

    logging.info(f"==================================================")
    logging.info(f"【パイプライン完了】 成功: {success_count} / {len(python_files)} 件")
    logging.info(f" 成果物出力先: {outputs_dir}")
    logging.info(f"==================================================")

if __name__ == "__main__":
    run_full_pipeline()