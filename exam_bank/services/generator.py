"""
문서 생성 엔진 v5.0
— 시험지 (2단 칼럼 + 디자인 박스) / 워크북 / 구문해석지 / 정답지
— Exam Maker PRO v2.1 의 디자인 테마 통합
"""
import os
import re
import time
import datetime
import hashlib
import shutil
import tempfile
import json
import platform

import docx
from docx.shared import Cm, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    from docx2pdf import convert as _convert_pdf
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

from exam_bank.constants import Theme
from exam_bank.services.parser import parse_exam_text

MAX_SAFE_PATH = 240

# ─── 유틸리티 ───

def _safe_basename(folder, stem, ext):
    full = os.path.join(folder, stem + ext)
    if len(full) <= MAX_SAFE_PATH:
        return stem
    h = hashlib.md5(stem.encode("utf-8")).hexdigest()[:6]
    cut_len = max(1, MAX_SAFE_PATH - (len(folder) + 1 + len(ext) + 8))
    return f"{stem[:cut_len]}_{h}"


def build_paths(target_dir, base_filename, prefix=""):
    stem_raw = f"{prefix}{base_filename}"
    safe_stem = _safe_basename(target_dir, stem_raw, ".docx")
    docx_path = os.path.join(target_dir, safe_stem + ".docx")
    pdf_path = os.path.join(target_dir, safe_stem + ".pdf")
    return docx_path, pdf_path


def convert_to_pdf_safe(docx_path, pdf_path):
    time.sleep(0.5)
    if platform.system() == "Windows":
        try:
            import win32com.client
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
            doc.Close()
            word.Quit()
            return True, "성공"
        except Exception as e:
            if PDF_AVAILABLE:
                try:
                    _convert_pdf(docx_path, pdf_path)
                    return True, "성공"
                except Exception:
                    pass
            return False, f"PDF 변환 오류: {e}"
    else:
        if not PDF_AVAILABLE:
            return False, "docx2pdf 미설치"
        try:
            _convert_pdf(docx_path, pdf_path)
            return True, "성공"
        except Exception as e:
            return False, str(e)


def convert_docx_pairs_to_pdf_checked(pairs):
    """여러 docx를 일괄 PDF 변환"""
    results = []
    for docx_path, pdf_path in pairs:
        if not os.path.exists(docx_path):
            continue
        ok, msg = convert_to_pdf_safe(docx_path, pdf_path)
        results.append((ok, msg))
    return results


# ─── 문서 스타일 헬퍼 ───

def _apply_doc_style(doc, font_name, font_size):
    style = doc.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(int(font_size))
    style.font.color.rgb = RGBColor.from_string(Theme.DARK_TEXT)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    section = doc.sections[0]
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)


def _set_two_columns(section):
    try:
        sectPr = section._sectPr
        cols = sectPr.xpath("./w:cols")
        cols_elem = cols[0] if cols else OxmlElement("w:cols")
        cols_elem.set(qn("w:num"), "2")
        cols_elem.set(qn("w:space"), "600")
        cols_elem.set(qn("w:sep"), "1")
        if not cols:
            sectPr.append(cols_elem)
    except Exception:
        pass


def set_cell_background(cell, fill_color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tcPr.append(shd)


def set_cell_borders(cell, color="B0C4DE", size="4"):
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), size)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)


def set_cell_vertical_alignment(cell, align="center"):
    tcPr = cell._element.get_or_add_tcPr()
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), align)
    tcPr.append(vAlign)


def set_paragraph_border_bottom(paragraph, color="2E86C1", size="6", space="3"):
    pPr = paragraph._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), size)
    bottom.set(qn('w:space'), space)
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_formatted_text(paragraph, text, font_size=None, color=None):
    pattern = r'([①-⑳ⓐ-ⓩⒶ-Ⓩ]|\([A-Ea-e]\))'
    tokens = re.split(pattern, text)
    for token in tokens:
        if not token:
            continue
        run = paragraph.add_run(token)
        if font_size:
            run.font.size = Pt(font_size)
        if color:
            run.font.color.rgb = RGBColor.from_string(color)
        if re.match(pattern, token):
            run.bold = True


def add_header_banner(doc, title_text, subtitle_text="", logo_path="", font_name="맑은 고딕"):
    if logo_path and os.path.exists(logo_path):
        logo_p = doc.add_paragraph()
        logo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            logo_p.add_run().add_picture(logo_path, height=Cm(2.5))
        except Exception:
            pass
        logo_p.paragraph_format.space_after = Pt(4)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(6)
    title_p.paragraph_format.space_after = Pt(2)
    title_run = title_p.add_run(title_text)
    title_run.bold = True
    title_run.font.size = Pt(18)
    title_run.font.color.rgb = RGBColor.from_string(Theme.PRIMARY)
    title_run.font.name = font_name
    title_run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)

    if subtitle_text:
        sub_p = doc.add_paragraph()
        sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = sub_p.add_run(subtitle_text)
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor.from_string(Theme.GRAY_TEXT)

    sep_p = doc.add_paragraph()
    sep_p.paragraph_format.space_before = Pt(0)
    sep_p.paragraph_format.space_after = Pt(8)
    set_paragraph_border_bottom(sep_p, color=Theme.SECONDARY, size="8")
    return title_p


def _set_cell_margins(cell, top='100', bottom='100', left='140', right='140'):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        elem = OxmlElement(f'w:{side}')
        elem.set(qn('w:w'), val)
        elem.set(qn('w:type'), 'dxa')
        tcMar.append(elem)
    tcPr.append(tcMar)


def create_styled_passage_table(doc, passage_text):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_background(cell, Theme.PASSAGE_BG)
    set_cell_borders(cell, color=Theme.TABLE_BORDER, size="6")
    set_cell_vertical_alignment(cell)
    _set_cell_margins(cell)
    cell.text = ""
    add_formatted_text(cell.paragraphs[0], passage_text, font_size=9)
    for p in cell.paragraphs:
        p.paragraph_format.line_spacing = 1.3
    return table


def create_styled_box(doc, label, content, bg_color, border_color, label_color_hex):
    label_p = doc.add_paragraph()
    label_p.paragraph_format.space_before = Pt(4)
    label_p.paragraph_format.space_after = Pt(2)
    label_run = label_p.add_run(f"  {label}")
    label_run.bold = True
    label_run.font.size = Pt(9)
    label_run.font.color.rgb = RGBColor.from_string(label_color_hex)
    if not content.strip():
        return
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_background(cell, bg_color)
    set_cell_borders(cell, color=border_color, size="4")
    _set_cell_margins(cell, '80', '80', '120', '120')
    cell.text = ""
    add_formatted_text(cell.paragraphs[0], content, font_size=9)
    return table


# ─── 시험지 생성 (Exam Maker PRO v2.1 디자인) ───

def create_full_exam_docx(target_dir, filename, raw_text, font_name, font_size,
                          logo_path="", display_title="실전 모의고사"):
    exam_data = parse_exam_text(raw_text)
    if not exam_data:
        raise ValueError("분석할 지문과 문제를 찾을 수 없습니다.")

    doc = docx.Document()
    _apply_doc_style(doc, font_name, font_size)

    today_display = datetime.datetime.now().strftime("%Y년 %m월 %d일")
    add_header_banner(doc, display_title,
                      f"📅 {today_display}  |  이름: ________________",
                      logo_path, font_name)

    new_section = doc.add_section(WD_SECTION.CONTINUOUS)
    _set_two_columns(new_section)

    for item in exam_data:
        if item["passage"]:
            p_intro = doc.add_paragraph()
            p_intro.paragraph_format.space_before = Pt(10)
            intro_run = p_intro.add_run("※ 다음 글을 읽고 물음에 답하시오.")
            intro_run.bold = True
            intro_run.font.size = Pt(int(font_size))
            intro_run.font.color.rgb = RGBColor.from_string(Theme.SECONDARY)
            create_styled_passage_table(doc, item["passage"])
            doc.add_paragraph().paragraph_format.space_after = Pt(4)

        for q in item["questions"]:
            q_text = q["text"]
            special_pattern = r'(<조건>|\[조건\]|\[우리말\]|\[보기\]|\[정답\]\s*:?|정답\s*:)'
            special_matches = list(re.finditer(special_pattern, q_text))
            main_q = q_text
            special_blocks = []

            if special_matches:
                main_q = q_text[:special_matches[0].start()].strip()
                for idx_m, sm in enumerate(special_matches):
                    label = sm.group(0).strip().rstrip(':').strip()
                    start = sm.end()
                    end = special_matches[idx_m + 1].start() if idx_m + 1 < len(special_matches) else len(q_text)
                    content = q_text[start:end].strip()
                    special_blocks.append((label, content))
            else:
                ans_match = re.search(r'(\[정답\]\s*:|정답\s*:)', q_text)
                if ans_match:
                    main_q = q_text[:ans_match.start()].strip()
                    special_blocks.append(("[정답]", q_text[ans_match.end():].strip()))

            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            num_run = p.add_run(f"{q['num']}. ")
            num_run.bold = True
            num_run.font.size = Pt(int(font_size) + 1)
            num_run.font.color.rgb = RGBColor.from_string(Theme.PRIMARY)
            add_formatted_text(p, main_q)

            for label, content in special_blocks:
                if '<조건>' in label or '[조건]' in label:
                    clean_c = re.sub(r'^(<조건>|\[조건\])\s*\n?', '', content).strip()
                    create_styled_box(doc, "<조건>", clean_c,
                                      Theme.CONDITION_BG, "E6C84C", "CC8800")
                elif '[우리말]' in label:
                    create_styled_box(doc, "[우리말]", content,
                                      "F5F0FF", "B39DDB", "5E35B1")
                elif '[보기]' in label:
                    create_styled_box(doc, "[보기]", content,
                                      "EFF8FF", "7FAFD4", "2E6DA4")
                elif '[정답]' in label or '정답' in label:
                    ans_p = doc.add_paragraph()
                    ans_p.paragraph_format.space_before = Pt(4)
                    ans_run = ans_p.add_run(f"[정답]: {content}" if content else "[정답]:")
                    ans_run.font.size = Pt(int(font_size) - 1)
                    ans_run.font.color.rgb = RGBColor.from_string(Theme.GRAY_TEXT)

            if q["choices"] and any(q["choices"]):
                c_type = q.get("c_type", "num")
                prefix = ["(A)", "(B)", "(C)", "(D)", "(E)"] if c_type == "alpha" else ["①", "②", "③", "④", "⑤"]
                for i in range(5):
                    if i < len(q['choices']) and q['choices'][i]:
                        cp = doc.add_paragraph()
                        cp.paragraph_format.left_indent = Cm(1.3)
                        cp.paragraph_format.first_line_indent = Cm(-0.8)
                        cp.paragraph_format.space_before = Pt(1)
                        cp.paragraph_format.space_after = Pt(1)
                        marker_run = cp.add_run(f"{prefix[i]}  ")
                        marker_run.bold = True
                        marker_run.font.size = Pt(int(font_size))
                        add_formatted_text(cp, q['choices'][i])

            spacer = doc.add_paragraph()
            spacer.paragraph_format.space_after = Pt(6)

    docx_path, pdf_path = build_paths(target_dir, filename, prefix="")
    doc.save(docx_path)
    return docx_path, pdf_path


# ─── 정답지 생성 ───

def create_answer_sheet_docx(target_dir, filename, raw_ans_text, font_name, font_size,
                             logo_path="", display_title="실전 모의고사"):
    doc = docx.Document()
    _apply_doc_style(doc, font_name, font_size)
    add_header_banner(doc, "정답 및 해설", display_title, logo_path, font_name)

    new_section = doc.add_section(WD_SECTION.CONTINUOUS)
    _set_two_columns(new_section)

    ans_m = re.findall(
        r'\b(\d+)(?:번|[\)\.])?(?:\s*정답\s*:)?\s*[\(\[]?([①②③④⑤A-Ea-e])[\)\]]?',
        raw_ans_text
    )

    if ans_m:
        label_p = doc.add_paragraph()
        label_run = label_p.add_run("  ✅ 빠른 정답 채점표")
        label_run.bold = True
        label_run.font.size = Pt(11)
        label_run.font.color.rgb = RGBColor.from_string(Theme.PRIMARY)
        label_p.paragraph_format.space_after = Pt(6)

        table = doc.add_table(rows=1, cols=10)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr = table.rows[0].cells
        for i in range(5):
            hdr[i*2].text = '문항'
            hdr[i*2+1].text = '정답'
            for c in (hdr[i*2], hdr[i*2+1]):
                set_cell_background(c, Theme.HEADER_BG)
                set_cell_borders(c, color=Theme.TABLE_BORDER)
                set_cell_vertical_alignment(c)
                for p in c.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for r in p.runs:
                        r.bold = True
                        r.font.size = Pt(8)
                        r.font.color.rgb = RGBColor.from_string(Theme.PRIMARY)

        row_cells = None
        for i, (q_n, ans) in enumerate(ans_m):
            if i % 5 == 0:
                row_cells = table.add_row().cells
            row_cells[(i%5)*2].text = q_n
            row_cells[(i%5)*2+1].text = ans.upper()
            set_cell_background(row_cells[(i%5)*2], Theme.LIGHT_BG)
            set_cell_borders(row_cells[(i%5)*2], color=Theme.TABLE_BORDER)
            set_cell_borders(row_cells[(i%5)*2+1], color=Theme.TABLE_BORDER)
            for p in (row_cells[(i%5)*2].paragraphs[0], row_cells[(i%5)*2+1].paragraphs[0]):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph()
        sep_p = doc.add_paragraph()
        set_paragraph_border_bottom(sep_p, color=Theme.TABLE_BORDER, size="4")
        doc.add_paragraph()

    for line in raw_ans_text.split("\n"):
        stripped = line.strip()
        if not stripped:
            continue
        num_m = re.match(r'^(\d+)[\.\)번]', stripped)
        p = doc.add_paragraph()
        if num_m:
            p.paragraph_format.space_before = Pt(6)
            num_run = p.add_run(stripped[:num_m.end()] + " ")
            num_run.bold = True
            num_run.font.color.rgb = RGBColor.from_string(Theme.SECONDARY)
            p.add_run(stripped[num_m.end():].strip())
        else:
            p.add_run(stripped)

    docx_path, pdf_path = build_paths(target_dir, filename, prefix="답안지_")
    doc.save(docx_path)
    return docx_path, pdf_path


# ─── 워크북 생성 ───

def create_workbook_docx(target_dir, filename, raw_text, font_name, font_size,
                         logo_path="", display_title="실전 모의고사"):
    exam_data = parse_exam_text(raw_text)
    if not exam_data:
        raise ValueError("분석할 지문을 찾을 수 없습니다.")

    doc = docx.Document()
    _apply_doc_style(doc, font_name, font_size)

    all_passages = []
    for item in exam_data:
        if item["passage"]:
            sub_blocks = re.split(
                r'(?=(?:\n|^)\s*(?:\d+[\.\)]\s*)?(?:\[?지문\s*\d+\]?|\[Questions?\s*\d+))',
                item["passage"], flags=re.IGNORECASE
            )
            for sb in sub_blocks:
                if sb.strip():
                    all_passages.append(sb.strip())
    if not all_passages:
        raise ValueError("워크북으로 변환할 텍스트가 부족합니다.")

    for p_idx, p_text in enumerate(all_passages):
        if p_idx > 0:
            doc.add_page_break()
        add_header_banner(doc, "본문 해석 연습 워크북",
                          f"{display_title}  ·  지문 {p_idx+1}", logo_path, font_name)
        inst = doc.add_paragraph()
        inst.paragraph_format.space_before = Pt(6)
        inst_run = inst.add_run("✏️ 문장 전체의 자연스러운 해석을 써 보세요.")
        inst_run.bold = True
        inst_run.font.color.rgb = RGBColor.from_string(Theme.SECONDARY)
        inst_run.font.size = Pt(10)
        doc.add_paragraph()

        clean_passage = re.sub(
            r'^(?:\d+[\.\)]\s*)?(?:\[?지문\s*\d+\]?|\[Questions?\s*\d+.*?\]|다음\s*글을\s*읽고.*?)\s*',
            '', p_text, flags=re.IGNORECASE
        )
        clean_passage = clean_passage.replace('\n', ' ')
        sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', clean_passage) if s.strip()]

        for s_idx, sent in enumerate(sentences, 1):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            num_run = p.add_run(f"  {s_idx}.  ")
            num_run.bold = True
            num_run.font.color.rgb = RGBColor.from_string(Theme.PRIMARY)
            num_run.font.size = Pt(int(font_size) + 1)
            add_formatted_text(p, sent)
            for _ in range(2):
                blank = doc.add_paragraph()
                blank.paragraph_format.line_spacing = 1.8
                blank.paragraph_format.left_indent = Cm(0.5)
                blank_run = blank.add_run("_" * 75)
                blank_run.font.color.rgb = RGBColor(210, 210, 210)
                blank_run.font.size = Pt(9)

    docx_path, pdf_path = build_paths(target_dir, filename, prefix="워크북_")
    doc.save(docx_path)
    return docx_path, pdf_path


# ─── 구문 해석지 생성 ───

def create_syntax_workbook_docx(target_dir, filename, raw_text, font_name, font_size,
                                logo_path="", display_title="실전 모의고사"):
    doc = docx.Document()
    _apply_doc_style(doc, font_name, font_size)
    add_header_banner(doc, "구문 및 단어 집중 해석지", display_title, logo_path, font_name)

    inst = doc.add_paragraph()
    inst.paragraph_format.space_before = Pt(6)
    inst_run = inst.add_run("🍀 다음 문장을 읽고, 타겟 단어/구문에 유의하여 자연스럽게 해석해 보세요.")
    inst_run.bold = True
    inst_run.font.color.rgb = RGBColor.from_string(Theme.GREEN)
    inst_run.font.size = Pt(10)
    doc.add_paragraph()

    raw_text = re.sub(r'\(?해석\)?\s*:', '', raw_text)
    raw_text = raw_text.replace('**', '')
    blocks = re.split(r'\n(?=\s*\d+[\.\)]\s+)', '\n' + raw_text)

    q_num_auto = 1
    for block in blocks:
        if not block.strip():
            continue
        m = re.match(r'^\s*(\d+)[\.\)]\s+(.*)', block, re.DOTALL)
        content = m.group(2) if m else block.strip()
        content = re.sub(r'\s*\n\s*', ' ', content).strip()
        if not content:
            continue

        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        run_num = p.add_run(f" [{str(q_num_auto).zfill(2)}]  ")
        run_num.bold = True
        run_num.font.color.rgb = RGBColor.from_string(Theme.GREEN)
        run_num.font.size = Pt(int(font_size) + 1)
        run_text = p.add_run(content)
        run_text.font.color.rgb = RGBColor.from_string(Theme.DARK_TEXT)
        run_text.font.size = Pt(int(font_size) + 1)

        for _ in range(2):
            blank = doc.add_paragraph()
            blank.paragraph_format.line_spacing = 1.8
            blank.paragraph_format.left_indent = Cm(0.5)
            blank_run = blank.add_run("_" * 80)
            blank_run.font.color.rgb = RGBColor(200, 200, 200)
            blank_run.font.size = Pt(9)

        q_num_auto += 1

    docx_path, pdf_path = build_paths(target_dir, filename, prefix="구문해석_")
    doc.save(docx_path)
    return docx_path, pdf_path


# ─── 기존 호환 함수 (bank_tab에서 사용) ───

def create_exam_files(target_dir, prefix, exam_data, cfg,
                      is_workbook=False, output_format="Word + PDF",
                      student_name="", set_label=""):
    """bank_tab.py 에서 장바구니 기반 시험지 생성 시 호출"""
    ts = datetime.datetime.now().strftime("%m%d_%H%M")
    label = f"_{set_label}" if set_label else ""
    font_name = cfg.get("font_name", "맑은 고딕")
    font_size = cfg.get("font_size", "10")
    logo_path = cfg.get("last_logo", "")

    doc_q = docx.Document()
    title_prefix = "워크북" if is_workbook else "시험지"
    title_text = f"실전 대비 맞춤형 {title_prefix}"
    if student_name:
        title_text += f" [{student_name}]"

    _apply_doc_style(doc_q, font_name, font_size)
    add_header_banner(doc_q, title_text, "", logo_path, font_name)

    if not is_workbook:
        new_sec = doc_q.add_section(WD_SECTION.CONTINUOUS)
        _set_two_columns(new_sec)

    g_idx = 1
    for idx, item in enumerate(exam_data, 1):
        if not item.get("passage"):
            continue
        item["q_mappings"] = []

        if is_workbook:
            h = doc_q.add_paragraph()
            h.add_run(f"■ 지문 {idx}  ").bold = True
            h.add_run(f"[{item.get('info', '')}]").font.color.rgb = RGBColor(100, 100, 100)
            sentences = [s.strip() for s in re.split(r"(?<=[.!?])\s+", item["passage"].strip()) if s.strip()]
            for s_i, sent in enumerate(sentences, 1):
                p = doc_q.add_paragraph()
                p.paragraph_format.space_before = Pt(10)
                p.add_run(f"{s_i}. {sent}").bold = True
                for _ in range(2):
                    bl = doc_q.add_paragraph()
                    bl.paragraph_format.space_after = Pt(0)
                    bl.add_run("_" * 74).font.color.rgb = RGBColor(220, 220, 220)
            doc_q.add_paragraph("\n")
        else:
            intro = doc_q.add_paragraph()
            intro_run = intro.add_run("※ 다음 글을 읽고 물음에 답하시오.")
            intro_run.bold = True
            intro_run.font.color.rgb = RGBColor.from_string(Theme.SECONDARY)

            create_styled_passage_table(doc_q, item["passage"])
            doc_q.add_paragraph()

            for q in item.get("questions", []):
                p = doc_q.add_paragraph()
                num_run = p.add_run(f"{g_idx}. ")
                num_run.bold = True
                num_run.font.color.rgb = RGBColor.from_string(Theme.PRIMARY)
                p.add_run(q.get("text", q.get("content", "")))

                choices = q.get("choices", [])
                if choices and any(c for c in choices if c):
                    if isinstance(choices, str):
                        try:
                            choices = json.loads(choices)
                        except json.JSONDecodeError:
                            choices = []
                    for ch in choices:
                        if ch:
                            cp = doc_q.add_paragraph()
                            cp.paragraph_format.left_indent = Cm(1.3)
                            cp.paragraph_format.first_line_indent = Cm(-0.8)
                            add_formatted_text(cp, ch)

                doc_q.add_paragraph()
                item["q_mappings"].append({
                    "new_num": g_idx,
                    "orig_num": q.get("num", q.get("q_num", "-")),
                    "preview": q.get("text", q.get("content", ""))[:25].replace("\n", " ") + "...",
                })
                g_idx += 1

    q_docx = os.path.normpath(os.path.join(target_dir, f"{prefix}_문제지_{ts}{label}.docx"))
    doc_q.save(q_docx)

    # 정답지
    doc_a = docx.Document()
    _apply_doc_style(doc_a, font_name, font_size)
    add_header_banner(doc_a, f"맞춤형 {title_prefix} — 정답 및 해설", "", logo_path, font_name)

    for idx, item in enumerate(exam_data, 1):
        h = doc_a.add_paragraph()
        h.add_run(f"■ [지문 {idx}] 해설\n").bold = True
        if not is_workbook and item.get("q_mappings"):
            pm = doc_a.add_paragraph()
            rt = pm.add_run("[※ 문항 번호 매칭표]\n")
            rt.bold = True
            rt.font.color.rgb = RGBColor(0, 112, 192)
            lines = ""
            for m in item["q_mappings"]:
                orig = f"{m['orig_num']}번" if m["orig_num"] != "-" else "서술형"
                lines += f"▶ 시험지 {m['new_num']}번 = 본문 해설 {orig} 참조 | {m['preview']}\n"
            rm = pm.add_run(lines)
            rm.font.color.rgb = RGBColor(100, 100, 100)
            rm.font.size = Pt(9)
            doc_a.add_paragraph()
        body = doc_a.add_paragraph()
        ans = (item.get("answer_text") or "").strip()
        body.add_run(ans if ans else "(입력된 정답/해설이 없습니다.)")
        doc_a.add_paragraph("\n" + "=" * 50 + "\n")

    a_docx = os.path.normpath(os.path.join(target_dir, f"{prefix}_정답지_{ts}{label}.docx"))
    doc_a.save(a_docx)

    msg = "Word 파일(문제/답안) 분리 생성 완료!"
    if output_format in ("PDF만", "Word + PDF"):
        q_pdf = q_docx.replace(".docx", ".pdf")
        a_pdf = a_docx.replace(".docx", ".pdf")
        q_ok, _ = convert_to_pdf_safe(q_docx, q_pdf)
        a_ok, _ = convert_to_pdf_safe(a_docx, a_pdf)
        if q_ok and a_ok:
            msg = "Word + PDF 분리 생성 완료!"
            if output_format == "PDF만":
                for f in (q_docx, a_docx):
                    try:
                        os.remove(f)
                    except OSError:
                        pass
                msg = "PDF 분리 생성 완료!"
        else:
            msg = "Word 생성 완료. PDF 변환 실패 (Word가 필요합니다)"
    return msg
