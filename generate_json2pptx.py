import os
import sys
import json
import re
import ast
import time
import logging
from pathlib import Path
from typing import List, Optional, Union, Any
from dotenv import load_dotenv
from pydantic import BaseModel, Field

# python-pptx モジュール
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# ---------------------------------------------------------
# 1. Pydantic 構造化データ定義 (柔軟なスキーマ対応)
# ---------------------------------------------------------

class FlowStepDetail(BaseModel):
    title: str = Field(default="", description="ステップのタイトル")
    details: List[str] = Field(default_factory=list, description="ステップの具体例・内訳項目")

class BusinessSummary(BaseModel):
    purpose_and_value: str = Field(default="", description="プログラムの目的と創出するビジネス価値")
    target_user: str = Field(default="", description="想定利用者")
    flow_step_details: List[FlowStepDetail] = Field(default_factory=list, description="詳細内訳付き実行フローステップ一覧")

class CodeSectionSlide(BaseModel):
    category_title: str = Field(default="主要実行ルーチン・制御ロジック", description="コンポーネント分類")
    items_left: List[str] = Field(default_factory=list, description="左側: 関数名とその役割解説")
    learnings_right: List[str] = Field(default_factory=list, description="右側: レビュー視点・技術的学び")

class TableDictItem(BaseModel):
    term: str = Field(default="", description="用語・機能名")
    library_name: str = Field(default="", description="使用ライブラリ")
    explanation: str = Field(default="", description="辞書的な解説・適用目的")

class ProcessingStep(BaseModel):
    step_number: int = Field(default=1)
    phase_title: str = Field(default="")
    process_detail: str = Field(default="")

class SectionDetail(BaseModel):
    summary_plain: str = Field(default="")
    technical_point: str = Field(default="")

class LearningPoint(BaseModel):
    topic: str = Field(default="")
    understanding_level: str = Field(default="")

class DictItem(BaseModel):
    category: str = Field(default="")
    code_expression: str = Field(default="")
    explanation: str = Field(default="")

class CodeAnalysisResult(BaseModel):
    program_name: str
    business_summary: BusinessSummary
    main_flow: List[ProcessingStep] = Field(default_factory=list)
    sections: List[Union[CodeSectionSlide, SectionDetail]] = Field(default_factory=list)
    learning_points: List[LearningPoint] = Field(default_factory=list)
    dictionary_table: List[TableDictItem] = Field(default_factory=list)
    dictionary_items: List[DictItem] = Field(default_factory=list)

    def get_normalized_flow(self) -> List[FlowStepDetail]:
        """異なるJSONスキーマの実行フローを標準化"""
        if self.business_summary and self.business_summary.flow_step_details:
            return self.business_summary.flow_step_details
        
        normalized = []
        for step in self.main_flow:
            normalized.append(
                FlowStepDetail(
                    title=f"{step.phase_title}",
                    details=[step.process_detail] if step.process_detail else []
                )
            )
        return normalized

    def get_normalized_sections(self) -> CodeSectionSlide:
        """異なるJSONスキーマの実装解説・学習内容を標準化（見出しと詳細を完全に分離）"""
        items_left = []
        learnings_right = []

        for sec in self.sections:
            if isinstance(sec, CodeSectionSlide):
                items_left.extend(sec.items_left)
                learnings_right.extend(sec.learnings_right)
            elif isinstance(sec, SectionDetail):
                raw_title = getattr(sec, "title", "") or sec.summary_plain
                
                if "。" in raw_title:
                    short_title = raw_title.split("。")[0]
                elif len(raw_title) > 25:
                    short_title = raw_title[:25] + "…"
                else:
                    short_title = raw_title

                detail = sec.technical_point if sec.technical_point else sec.summary_plain
                items_left.append(f"{short_title.strip()}: {detail.strip()}")

        if self.learning_points:
            for lp in self.learning_points:
                topic_clean = lp.topic.strip().strip("【】")
                learnings_right.append(f"【{topic_clean}】: {lp.understanding_level.strip()}")

        category_title = "主要実行ルーチン・制御ロジック"
        if self.sections and isinstance(self.sections[0], CodeSectionSlide):
            category_title = self.sections[0].category_title

        return CodeSectionSlide(
            category_title=category_title,
            items_left=items_left,
            learnings_right=learnings_right
        )

    def get_normalized_dictionary(self) -> List[TableDictItem]:
        """異なるJSONスキーマの辞書データを標準化"""
        if self.dictionary_table:
            return self.dictionary_table
        
        normalized = []
        for item in self.dictionary_items:
            normalized.append(
                TableDictItem(
                    term=item.code_expression,
                    library_name=item.category,
                    explanation=item.explanation
                )
            )
        return normalized


# ---------------------------------------------------------
# 2. 静的コード解析 (AST) による動的解読 (単体実行フォールバック用)
# ---------------------------------------------------------

def _create_fallback_result(file_path: Path) -> CodeAnalysisResult:
    """ASTベースの動的解読フォールバック処理"""
    try:
        source_code = file_path.read_text(encoding="utf-8")
        tree = ast.parse(source_code)
    except Exception as e:
        logging.error(f"ファイル解析エラー: {file_path} - {e}")
        tree = ast.Module(body=[], type_ignores=[])

    pydantic_classes = []
    other_classes = []
    functions = []

    for node in tree.body:
        if isinstance(node, ast.ClassDef):
            is_pydantic = any(
                (isinstance(base, ast.Name) and base.id == "BaseModel") for base in node.bases
            )
            if is_pydantic:
                pydantic_classes.append(node.name)
            else:
                other_classes.append(node.name)
        elif isinstance(node, ast.FunctionDef):
            doc = ast.get_docstring(node)
            first_line = doc.split("\n")[0] if doc else "主要処理ルーチン"
            functions.append((node.name, first_line))

    imports = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                imports.append(alias.name)
        elif isinstance(node, ast.ImportFrom) and node.module:
            imports.append(node.module)
    imports = sorted(list(set(imports)))

    top_imports = [imp for imp in imports if imp in ["pandas", "matplotlib", "pptx", "seaborn", "tenacity", "pydantic", "pathlib", "os", "logging"]]
    imports_str = ", ".join(top_imports[:5]) if top_imports else "標準・外部拡張モジュール"

    total_classes = len(pydantic_classes) + len(other_classes)
    class_detail_str = f"Pydanticモデル ({', '.join(pydantic_classes)})" if pydantic_classes else "クラス定義なし (手続き型・関数主導の設計)"

    func_names = [f"{name}()" for name, _ in functions[:4]]
    func_detail_str = f"主要処理: {', '.join(func_names)}" if func_names else "メインスクリプト制御"

    flow_step_details = [
        FlowStepDetail(
            title=f"主要モジュール・依存ライブラリのロード (計{len(imports)}件)",
            details=[f"主要依存: {imports_str} 等"]
        ),
        FlowStepDetail(
            title=f"データ構造定義およびオブジェクトクラスの初期化 (計{total_classes}クラス)",
            details=[class_detail_str]
        ),
        FlowStepDetail(
            title=f"主要関数ルーチンによるデータ処理制御 (計{len(functions)}ルーチン)",
            details=[func_detail_str]
        ),
        FlowStepDetail(
            title="解析・レポート出力結果の生成および保存",
            details=["生成された分析図および考察サマリーをPowerPoint/Excel形式で出力"]
        )
    ]

    sections = []

    if pydantic_classes:
        sections.append(
            CodeSectionSlide(
                category_title="データ構造定義 (Pydanticモデル群)",
                items_left=[f"{cls}: 入出力データの型検証および構造化データの保持" for cls in pydantic_classes],
                learnings_right=[
                    "BaseModel利用: 堅牢な型安全性の確保",
                    "Structured Output: スキーマ準拠の保証"
                ]
            )
        )

    if functions:
        sections.append(
            CodeSectionSlide(
                category_title="主要実行ルーチン・制御ロジック",
                items_left=[f"{name}(): {doc}" for name, doc in functions[:5]],
                learnings_right=[
                    "モジュール化設計: 処理ステップの明確化と可読性向上",
                    "例外処理実装: 堅牢なエラーハンドリング"
                ]
            )
        )

    dict_table = []
    lib_map = {
        "calendar": ("monthrange()", "calendar", "月の日数や曜日計算を行う標準モジュール。"),
        "datetime": ("datetime / date", "datetime", "日付および時間の操作・フォーマット変換を行う標準モジュール。"),
        "glob": ("glob()", "glob", "ワイルドカードを使用したファイルパスのパターンマッチング検索機能。"),
        "io": ("StringIO / BytesIO", "io", "メモリ上でのインメモリデータストリーム入出力を扱う標準機能。"),
        "logging": ("getLogger / info", "logging", "プログラム実行時のイベント・エラー履歴を記録するログ機能。"),
        "matplotlib": ("pyplot / figure", "matplotlib", "2次元グラフやデータ可視化描画を行う基本ライブラリ。"),
        "numpy": ("ndarray / array()", "numpy", "大規模多次元配列や高度な数値計算を高速処理するライブラリ。"),
        "os": ("path / getenv()", "os", "OS依存の機能（環境変数・ファイルシステム操作等）を提供するモジュール。"),
        "pandas": ("DataFrame / Series", "pandas", "表形式データの操作・集計・データ分析を行うデータ解析モジュール。"),
        "pathlib": ("Path", "pathlib", "オブジェクト指向でファイルパスやディレクトリ操作を行う標準ライブラリ。"),
        "pptx": ("Presentation / Inches", "python-pptx", "PowerPoint（.pptx）ファイルの生成・自動編集を行う外部ライブラリ。"),
        "openpyxl": ("load_workbook()", "openpyxl", "Excelファイル（.xlsx/.xlsm）の読み込み・編集・保存を行う拡張ライブラリ。"),
        "tkinter": ("messagebox", "tkinter", "GUIメッセージボックスを表示してユーザーへ結果通知・エラー通知する標準機能。")
    }

    for lib in imports:
        if lib in lib_map:
            term, lib_name, exp = lib_map[lib]
            dict_table.append(TableDictItem(term=term, library_name=lib_name, explanation=exp))
        else:
            dict_table.append(TableDictItem(
                term="モジュール機能",
                library_name=lib,
                explanation=f"{lib} 機能を提供する標準/外部拡張ライブラリ。"
            ))

    return CodeAnalysisResult(
        program_name=file_path.name,
        business_summary=BusinessSummary(
            purpose_and_value=f"【AST解析】{file_path.name} の静的構造および依存関係の可視化",
            target_user="開発者 / 技術レビュー担当者",
            flow_step_details=flow_step_details
        ),
        sections=sections,
        dictionary_table=dict_table
    )


# ---------------------------------------------------------
# 3. PowerPoint 生成エンジン
# ---------------------------------------------------------

class PresentationGenerator:
    def __init__(self, template_path: Path):
        if not template_path.exists():
            raise FileNotFoundError(f"テンプレートが見つかりません: {template_path}")
        
        self.prs = Presentation(str(template_path))
        
        for i in range(len(self.prs.slides) - 1, -1, -1):
            rId = self.prs.slides._sldIdLst[i].rId
            self.prs.part.drop_rel(rId)
            del self.prs.slides._sldIdLst[i]

        self.layouts = {layout.name: layout for layout in self.prs.slide_layouts}

    def _get_layout(self, possible_names: list, default_index: int = 0):
        for name in possible_names:
            if name in self.layouts:
                return self.layouts[name]
        return self.prs.slide_layouts[default_index]

    def _add_page_number(self, slide, page_num: int):
        tx_box = slide.shapes.add_textbox(Inches(11.5), Inches(6.8), Inches(1.5), Inches(0.4))
        tf = tx_box.text_frame
        p = tf.paragraphs[0]
        p.text = f"- Page {page_num} -"
        p.font.size = Pt(11)
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = PP_ALIGN.RIGHT

    def create_title_slide(self, program_name: str) -> int:
        """1ページ目: 表紙"""
        layout = self._get_layout(["タイトル スライド", "Title Slide"], default_index=0)
        slide = self.prs.slides.add_slide(layout)
        page_num = 1

        title_text = f"コード解析・研修成果レポート\n【 {program_name} 】"
        info_text = "（例）業務DX推進部 / システム開発チーム\n氏 名：＊＊ ＊＊\n作成日：2026年10月"

        title_shape = None
        subtitle_shape = None

        for shape in slide.placeholders:
            ph_type = shape.placeholder_format.type
            if ph_type == 1 or shape.name.startswith("Title"):
                title_shape = shape
            elif ph_type in (2, 7) or shape.name.startswith("Sub") or shape.name.startswith("Text"):
                subtitle_shape = shape

        if title_shape and title_shape.has_text_frame:
            tf = title_shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = title_text
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = RGBColor(24, 43, 73)

        if subtitle_shape and subtitle_shape.has_text_frame:
            subtitle_shape.left = Inches(8.66)
            subtitle_shape.top = Inches(5.59)
            subtitle_shape.width = Inches(3.98)
            subtitle_shape.height = Inches(1.04)
            tf = subtitle_shape.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.TOP
            p = tf.paragraphs[0]
            p.text = info_text
            p.font.size = Pt(15)
            p.font.color.rgb = RGBColor(60, 60, 60)
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.3

        return page_num

    def create_summary_flow_slide(self, summary: BusinessSummary, flow_steps: List[FlowStepDetail], current_page: int) -> int:
        """2ページ目: 大見出し・中見出し・詳細の段階的な段分けインデント構造"""
        layout = self._get_layout(["1_セクション見出し", "Section Header", "タイトルとコンテンツ"], default_index=1)
        slide = self.prs.slides.add_slide(layout)
        page_num = current_page + 1

        title_shape = None
        content_shape = None

        for shape in slide.placeholders:
            ph_type = shape.placeholder_format.type
            if ph_type == 1 or shape.name.startswith("Title"):
                title_shape = shape
            elif ph_type in (2, 7) or shape.name.startswith("Text") or shape.name.startswith("Content"):
                content_shape = shape

        if title_shape and title_shape.has_text_frame:
            tf = title_shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = "■ 概要 & 全体実行フロー解説"
            p.font.size = Pt(24)
            p.font.bold = True
            p.font.color.rgb = RGBColor(24, 43, 73)

        if content_shape and content_shape.has_text_frame:
            tf = content_shape.text_frame
            tf.word_wrap = True
            
            p1 = tf.paragraphs[0]
            p1.text = "【解析目的 & 想定ユーザー】"
            p1.font.size = Pt(14)
            p1.font.bold = True
            p1.font.color.rgb = RGBColor(0, 102, 204)

            p2_1 = tf.add_paragraph()
            p2_1.text = f"    目的: {summary.purpose_and_value}"
            p2_1.font.size = Pt(11)
            p2_1.line_spacing = 1.2
            p2_1.space_before = Pt(3)

            p2_2 = tf.add_paragraph()
            p2_2.text = f"    対象: {summary.target_user}"
            p2_2.font.size = Pt(11)
            p2_2.line_spacing = 1.2
            p2_2.space_before = Pt(2)

            p3 = tf.add_paragraph()
            p3.text = "【プログラム全体実行フロー（処理ステップ・内訳）】"
            p3.font.size = Pt(14)
            p3.font.bold = True
            p3.font.color.rgb = RGBColor(0, 102, 204)
            p3.space_before = Pt(10)

            for i, step in enumerate(flow_steps, start=1):
                p_step = tf.add_paragraph()
                p_step.text = f"      {i}. {step.title}"
                p_step.font.size = Pt(11)
                p_step.font.bold = True
                p_step.space_before = Pt(4)

                for detail in step.details:
                    p_detail = tf.add_paragraph()
                    p_detail.text = f"            └ {detail}"
                    p_detail.font.size = Pt(10)
                    p_detail.font.color.rgb = RGBColor(80, 80, 80)
                    p_detail.line_spacing = 1.15
                    p_detail.space_before = Pt(1)

        self._add_page_number(slide, page_num)
        return page_num

    def create_comparison_section_slide(self, section: CodeSectionSlide, current_page: int) -> int:
        """3ページ目以降: 左右両枠ともに「タイトル（青太字）」と「└ 説明（グレー）」の2行分割表示"""
        layout = self._get_layout(["比較", "Comparison"], default_index=4)
        slide = self.prs.slides.add_slide(layout)
        page_num = current_page + 1

        title_shape = None
        body_shapes = []

        for shape in slide.placeholders:
            ph_type = shape.placeholder_format.type
            if ph_type == 1 or shape.name.startswith("Title"):
                title_shape = shape
            else:
                body_shapes.append(shape)

        if title_shape and title_shape.has_text_frame:
            tf = title_shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = f"■ {section.category_title}"
            p.font.size = Pt(22)
            p.font.bold = True
            p.font.color.rgb = RGBColor(24, 43, 73)

        left_header_shape = None
        left_body_shape = None
        right_header_shape = None
        right_body_shape = None

        if len(body_shapes) >= 4:
            left_group = [s for s in body_shapes if s.left < Inches(6.6)]
            right_group = [s for s in body_shapes if s.left >= Inches(6.6)]

            left_group.sort(key=lambda s: s.top)
            right_group.sort(key=lambda s: s.top)

            if len(left_group) >= 2:
                left_header_shape, left_body_shape = left_group[0], left_group[1]
            if len(right_group) >= 2:
                right_header_shape, right_body_shape = right_group[0], right_group[1]

        if not (left_header_shape and left_body_shape and right_header_shape and right_body_shape):
            for b_shape in body_shapes:
                try:
                    sp_element = b_shape._element
                    sp_element.getparent().remove(sp_element)
                except Exception:
                    pass

            left_header_shape = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(5.15), Inches(0.6))
            left_body_shape = slide.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(5.15), Inches(4.3))
            right_header_shape = slide.shapes.add_textbox(Inches(7.47), Inches(1.5), Inches(5.15), Inches(0.6))
            right_body_shape = slide.shapes.add_textbox(Inches(7.47), Inches(2.2), Inches(5.15), Inches(4.3))

        # (A) 左・濃い枠
        if left_header_shape and left_header_shape.has_text_frame:
            tf = left_header_shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = "【処理内容・実装解説】"
            p.font.size = Pt(15)
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)

        # (B) 左・薄い枠（処理内容・実装解説）
        if left_body_shape and left_body_shape.has_text_frame:
            tf = left_body_shape.text_frame
            tf.word_wrap = True
            tf.paragraphs[0].text = ""

            for idx, item in enumerate(section.items_left):
                if ":" in item:
                    func_name, desc = item.split(":", 1)
                elif "：" in item:
                    func_name, desc = item.split("：", 1)
                else:
                    func_name, desc = item, ""

                p_func = tf.add_paragraph() if (idx > 0 or tf.paragraphs[0].text) else tf.paragraphs[0]
                p_func.text = f"■ {func_name.strip()}"
                p_func.font.size = Pt(12)
                p_func.font.bold = True
                p_func.font.color.rgb = RGBColor(0, 102, 204)
                p_func.space_before = Pt(8 if idx > 0 else 2)

                if desc.strip():
                    p_desc = tf.add_paragraph()
                    p_desc.text = f"    └ {desc.strip()}"
                    p_desc.font.size = Pt(11)
                    p_desc.font.color.rgb = RGBColor(60, 60, 60)
                    p_desc.line_spacing = 1.15
                    p_desc.space_before = Pt(1)

        # (C) 右・濃い枠
        if right_header_shape and right_header_shape.has_text_frame:
            tf = right_header_shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = "【コードから得られる学習内容】"
            p.font.size = Pt(15)
            p.font.bold = True
            p.font.color.rgb = RGBColor(255, 255, 255)

        # (D) 右・薄い枠（コードから得られる学習内容: タイトルと説明文を└で2行分割）
        if right_body_shape and right_body_shape.has_text_frame:
            tf = right_body_shape.text_frame
            tf.word_wrap = True
            tf.paragraphs[0].text = ""

            for idx, item in enumerate(section.learnings_right):
                if ":" in item:
                    title_part, desc_part = item.split(":", 1)
                elif "：" in item:
                    title_part, desc_part = item.split("：", 1)
                else:
                    title_part, desc_part = item, ""

                p_title = tf.add_paragraph() if (idx > 0 or tf.paragraphs[0].text) else tf.paragraphs[0]
                p_title.text = f"■ {title_part.strip()}"
                p_title.font.size = Pt(12)
                p_title.font.bold = True
                p_title.font.color.rgb = RGBColor(0, 102, 204)
                p_title.space_before = Pt(8 if idx > 0 else 2)

                if desc_part.strip():
                    p_desc = tf.add_paragraph()
                    p_desc.text = f"    └ {desc_part.strip()}"
                    p_desc.font.size = Pt(11)
                    p_desc.font.color.rgb = RGBColor(60, 60, 60)
                    p_desc.line_spacing = 1.15
                    p_desc.space_before = Pt(1)

        self._add_page_number(slide, page_num)
        return page_num

    def create_appendix_section_slide(self, items: List[TableDictItem], current_page: int) -> int:
        """4ページ目: 表形式スライド"""
        layout = self._get_layout(["1_セクション見出し", "Section Header", "タイトルとコンテンツ"], default_index=2)
        slide = self.prs.slides.add_slide(layout)
        page_num = current_page + 1

        title_shape = None
        unused_placeholders = []

        for shape in slide.placeholders:
            ph_type = shape.placeholder_format.type
            if ph_type == 1 or shape.name.startswith("Title"):
                title_shape = shape
            else:
                unused_placeholders.append(shape)

        if title_shape and title_shape.has_text_frame:
            tf = title_shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = "■ Appendix: 主要ライブラリ・用語解説辞書"
            p.font.size = Pt(22)
            p.font.bold = True
            p.font.color.rgb = RGBColor(24, 43, 73)

        for ph in unused_placeholders:
            try:
                sp_element = ph._element
                sp_element.getparent().remove(sp_element)
            except Exception:
                pass

        rows = len(items) + 1
        table_shape = slide.shapes.add_table(rows, 3, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.0))
        table = table_shape.table

        table.columns[0].width = Inches(3.0)
        table.columns[1].width = Inches(2.2)
        table.columns[2].width = Inches(6.5)

        headers = ["用語・機能", "ライブラリ", "解説・利用目的"]
        for col_idx, h_text in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = h_text
            p = cell.text_frame.paragraphs[0]
            p.font.bold = True
            p.font.size = Pt(13)
            p.font.color.rgb = RGBColor(255, 255, 255)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(24, 43, 73)

        for row_idx, item in enumerate(items, start=1):
            cell_term = table.cell(row_idx, 0)
            cell_lib = table.cell(row_idx, 1)
            cell_exp = table.cell(row_idx, 2)

            cell_term.text = item.term
            cell_lib.text = item.library_name
            cell_exp.text = item.explanation

            for cell in (cell_term, cell_lib, cell_exp):
                p = cell.text_frame.paragraphs[0]
                p.font.size = Pt(11)

        self._add_page_number(slide, page_num)
        return page_num

    def save(self, output_path: Path):
        output_path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(output_path))


# ---------------------------------------------------------
# 4. パイプライン・呼び出し用エントリーポイント
# ---------------------------------------------------------

def render_pptx_from_analysis(analysis_obj: Any, template_path: Path, output_path: Path):
    """構造化データを元にPPTXをレンダリング保存する標準関数"""
    
    if not hasattr(analysis_obj, "get_normalized_flow"):
        if hasattr(analysis_obj, "model_dump"):
            analysis_obj = CodeAnalysisResult.model_validate(analysis_obj.model_dump())
        elif isinstance(analysis_obj, dict):
            analysis_obj = CodeAnalysisResult.model_validate(analysis_obj)
        elif isinstance(analysis_obj, str):
            analysis_obj = CodeAnalysisResult.model_validate_json(analysis_obj)

    gen = PresentationGenerator(template_path=template_path)

    flow_steps = analysis_obj.get_normalized_flow()
    section_data = analysis_obj.get_normalized_sections()
    dict_items = analysis_obj.get_normalized_dictionary()

    current_page = gen.create_title_slide(analysis_obj.program_name)
    current_page = gen.create_summary_flow_slide(analysis_obj.business_summary, flow_steps, current_page)
    current_page = gen.create_comparison_section_slide(section_data, current_page)

    if dict_items:
        current_page = gen.create_appendix_section_slide(dict_items, current_page)

    gen.save(output_path)

if __name__ == "__main__":
    base_dir = Path.cwd()
    inputs_dir = base_dir / "inputs"
    outputs_dir = base_dir / "outputs"
    template_file = base_dir / "templates" / "aspark_org.pptx"

    python_files = list(inputs_dir.glob("*.py"))
    for target_script in python_files:
        analysis_data = _create_fallback_result(target_script)
        out_path = outputs_dir / f"成果報告_コード解説_{target_script.name}.pptx"
        render_pptx_from_analysis(analysis_data, template_file, out_path)