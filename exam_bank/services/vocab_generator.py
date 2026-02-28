"""
단어 시험지 생성 엔진 v7.0
— FactoryVoca Pro 스타일 2단 레이아웃
— 영→한 / 한→영 모드, 정답지, 첨글자 힌트
— v7.0: 예문 시험지 (문장해석 / 빈칸영작 Cloze)
"""

import os
import random
import datetime

import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from exam_bank.constants import Theme
from exam_bank.services.docx_utils import (
    apply_doc_style, add_header_banner, build_paths,
    safe_save_docx, convert_to_pdf_safe, sanitize_filename,
    set_cell_borders, set_cell_background, set_cell_vertical_alignment,
    set_cell_margins, set_paragraph_border_bottom, set_cell_width,
)


def _make_answer_hint(word, show_first_letter=False):
    """첨글자 힌트 생성"""
    if not show_first_letter or not word:
        return ""
    return word[0] + "_" * (len(word) - 1)


def _build_vocab_table(doc, words, exam_type, show_answer=False,
                       show_first_letter=False, font_name="맑은 고딕"):
    """
    FactoryVoca 스타일 2단 단어 테이블 생성
    words: [{"english": ..., "korean": ..., "pos": ...}, ...]
    """
    total = len(words)
    if total == 0:
        doc.add_paragraph("(출제할 단어가 없습니다)")
        return

    half = (total + 1) // 2
    left_words = words[:half]
    right_words = words[half:]

    rows_needed = max(len(left_words), len(right_words))
    table = doc.add_table(rows=rows_needed + 1, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    header_labels = ["No.", "단어", "뜻"] if exam_type == "eng_to_kor" else ["No.", "뜻", "단어"]
    header_labels_full = header_labels + header_labels
    header_row = table.rows[0]
    for i, label in enumerate(header_labels_full):
        cell = header_row.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(label)
        run.bold = True
        run.font.size = Pt(9)
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
        run.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, Theme.PRIMARY)
        set_cell_borders(cell, color=Theme.TABLE_BORDER, size="4")
        set_cell_vertical_alignment(cell)

    for row_idx in range(rows_needed):
        row = table.rows[row_idx + 1]
        for side in range(2):
            word_list = left_words if side == 0 else right_words
            col_offset = side * 3
            if row_idx < len(word_list):
                w = word_list[row_idx]
                num = row_idx + 1 if side == 0 else half + row_idx + 1

                if exam_type == "eng_to_kor":
                    question_text = w.get("english", "")
                    pos_text = f" ({w.get('pos', '')})" if w.get("pos") else ""
                    question_text += pos_text
                    answer_text = w.get("korean", "")
                else:
                    question_text = w.get("korean", "")
                    answer_text = w.get("english", "")

                num_cell = row.cells[col_offset]
                num_cell.text = ""
                p = num_cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(str(num))
                run.font.size = Pt(9)
                run.font.name = font_name
                run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
                set_cell_borders(num_cell, color=Theme.TABLE_BORDER, size="2")
                set_cell_vertical_alignment(num_cell)

                q_cell = row.cells[col_offset + 1]
                q_cell.text = ""
                p = q_cell.paragraphs[0]
                run = p.add_run(question_text)
                run.font.size = Pt(9)
                run.font.name = font_name
                run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
                set_cell_borders(q_cell, color=Theme.TABLE_BORDER, size="2")
                set_cell_vertical_alignment(q_cell)
                set_cell_margins(q_cell, left='80', right='80')

                a_cell = row.cells[col_offset + 2]
                a_cell.text = ""
                p = a_cell.paragraphs[0]
                if show_answer:
                    run = p.add_run(answer_text)
                    run.font.size = Pt(9)
                    run.font.color.rgb = RGBColor.from_string(Theme.ACCENT)
                    run.font.name = font_name
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
                elif show_first_letter:
                    hint = _make_answer_hint(answer_text, True)
                    run = p.add_run(hint)
                    run.font.size = Pt(9)
                    run.font.color.rgb = RGBColor(180, 180, 180)
                    run.font.name = font_name
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
                set_cell_borders(a_cell, color=Theme.TABLE_BORDER, size="2")
                set_cell_vertical_alignment(a_cell)
                set_cell_margins(a_cell, left='80', right='80')
            else:
                for c in range(3):
                    cell = row.cells[col_offset + c]
                    cell.text = ""
                    set_cell_borders(cell, color=Theme.TABLE_BORDER, size="2")

    for row in table.rows:
        for side in range(2):
            offset = side * 3
            set_cell_width(row.cells[offset], 1.0)
            set_cell_width(row.cells[offset + 1], 3.8)
            set_cell_width(row.cells[offset + 2], 3.8)

    return table


# ═══════════════════════════════════════════
# v7.0: 예문 시험지 헬퍼
# ═══════════════════════════════════════════

def _make_cloze_sentence(sentence_en, target_form):
    """빈칸 영작: target_form을 밑줄로 대체"""
    import re
    if not target_form:
        return sentence_en
    pattern = re.compile(re.escape(target_form), re.IGNORECASE)
    blank = "_" * max(len(target_form), 8)
    result = pattern.sub(blank, sentence_en, count=1)
    return result


def _make_cloze_with_hint(sentence_en, target_form):
    """빈칸 + 첫글자 힌트"""
    import re
    if not target_form:
        return sentence_en
    pattern = re.compile(re.escape(target_form), re.IGNORECASE)
    m = pattern.search(sentence_en)
    if not m:
        return sentence_en
    matched = m.group()
    hint = matched[0] + "_" * max(len(target_form) - 1, 6)
    result = sentence_en[:m.start()] + hint + sentence_en[m.end():]
    return result


def _add_highlighted_sentence(paragraph, sentence_en, target_form, font_name,
                               font_size=9, is_answer=False):
    """정답지용: target_form을 볼드+빨간색으로 하이라이트"""
    import re
    if not is_answer or not target_form:
        run = paragraph.add_run(sentence_en)
        run.font.size = Pt(font_size)
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
        return

    pattern = re.compile(re.escape(target_form), re.IGNORECASE)
    last_end = 0
    for m in pattern.finditer(sentence_en):
        if m.start() > last_end:
            run = paragraph.add_run(sentence_en[last_end:m.start()])
            run.font.size = Pt(font_size)
            run.font.name = font_name
            run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
        run = paragraph.add_run(m.group())
        run.font.size = Pt(font_size)
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
        run.bold = True
        run.underline = True
        run.font.color.rgb = RGBColor.from_string(Theme.ACCENT)
        last_end = m.end()
    if last_end < len(sentence_en):
        run = paragraph.add_run(sentence_en[last_end:])
        run.font.size = Pt(font_size)
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)


def _build_sentence_table(doc, items, test_mode, show_answer=False,
                           show_hint=False, font_name="맑은 고딕"):
    """
    예문 시험지 테이블 (1단 레이아웃)
    test_mode: "translation" (문장해석) | "cloze" (빈칸영작)
    """
    if not items:
        doc.add_paragraph("(출제할 예문이 없습니다)")
        return

    table = doc.add_table(rows=len(items) + 1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    if test_mode == "translation":
        headers = ["No.", "English Sentence", "한국어 해석"]
    else:
        headers = ["No.", "문장 (빈칸 채우기)", "정답"]

    header_row = table.rows[0]
    for i, label in enumerate(headers):
        cell = header_row.cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(label)
        run.bold = True
        run.font.size = Pt(9)
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
        run.font.color.rgb = RGBColor(255, 255, 255)
        set_cell_background(cell, Theme.PRIMARY)
        set_cell_borders(cell, color=Theme.TABLE_BORDER, size="4")
        set_cell_vertical_alignment(cell)

    for idx, item in enumerate(items):
        row = table.rows[idx + 1]
        en = item.get("sentence_en", "")
        ko = item.get("sentence_ko", "")
        tgt = item.get("target_form", "")
        word_eng = item.get("word_english", "")

        num_cell = row.cells[0]
        num_cell.text = ""
        p = num_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(str(idx + 1))
        run.font.size = Pt(9)
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
        set_cell_borders(num_cell, color=Theme.TABLE_BORDER, size="2")
        set_cell_vertical_alignment(num_cell)

        q_cell = row.cells[1]
        q_cell.text = ""
        p = q_cell.paragraphs[0]

        if test_mode == "translation":
            _add_highlighted_sentence(p, en, tgt, font_name, is_answer=show_answer)
        else:
            run_ko = p.add_run(f"[{ko}]")
            run_ko.font.size = Pt(8)
            run_ko.font.name = font_name
            run_ko.font.color.rgb = RGBColor.from_string(Theme.SECONDARY)
            run_ko._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
            p.add_run("\n")
            if show_hint:
                cloze = _make_cloze_with_hint(en, tgt)
            else:
                cloze = _make_cloze_sentence(en, tgt)
            run_q = p.add_run(cloze)
            run_q.font.size = Pt(9)
            run_q.font.name = font_name
            run_q._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)

        set_cell_borders(q_cell, color=Theme.TABLE_BORDER, size="2")
        set_cell_vertical_alignment(q_cell)
        set_cell_margins(q_cell, left='100', right='100')

        a_cell = row.cells[2]
        a_cell.text = ""
        p = a_cell.paragraphs[0]

        if show_answer:
            if test_mode == "translation":
                run = p.add_run(ko)
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor.from_string(Theme.ACCENT)
            else:
                run = p.add_run(tgt or word_eng)
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor.from_string(Theme.ACCENT)
                run.bold = True
            run.font.name = font_name
            run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)

        set_cell_borders(a_cell, color=Theme.TABLE_BORDER, size="2")
        set_cell_vertical_alignment(a_cell)
        set_cell_margins(a_cell, left='80', right='80')

        if idx % 2 == 1:
            for c in range(3):
                set_cell_background(row.cells[c], Theme.LIGHT_BG)

    for row in table.rows:
        set_cell_width(row.cells[0], 1.0)
        set_cell_width(row.cells[1], 10.0)
        set_cell_width(row.cells[2], 5.5)

    return table


def create_vocab_test(target_dir, words, cfg,
                      exam_type="eng_to_kor",
                      student_name="",
                      show_answer=False,
                      show_first_letter=False,
                      shuffle=False,
                      custom_title="",
                      custom_filename="",
                      logo_path="",
                      output_format="Word + PDF",
                      unit_names=None):
    """단어 시험지 생성 메인 함수"""
    if not words:
        return [], "출제할 단어가 없습니다."

    os.makedirs(target_dir, exist_ok=True)

    font_name = cfg.get("font_name", "맑은 고딕")
    font_size = cfg.get("font_size", "10")
    if not logo_path:
        logo_path = cfg.get("last_logo", "")

    ts = datetime.datetime.now().strftime("%m%d_%H%M")
    mode_label = "영한" if exam_type == "eng_to_kor" else "한영"

    if custom_title:
        title_text = custom_title
    else:
        title_text = "Vocabulary Test"
        if student_name:
            title_text += f" [{student_name}]"

    if custom_filename:
        base_name = sanitize_filename(custom_filename)
    else:
        name_part = f"_{student_name}" if student_name else ""
        base_name = f"VocabTest_{mode_label}{name_part}_{ts}"

    subtitle_parts = []
    if unit_names:
        if len(unit_names) <= 3:
            subtitle_parts.append(" / ".join(unit_names))
        else:
            subtitle_parts.append(f"{unit_names[0]} ~ {unit_names[-1]} ({len(unit_names)}개 단원)")
    subtitle_parts.append(f"{mode_label} | {len(words)}문항")
    subtitle_text = "  |  ".join(subtitle_parts)

    if shuffle:
        words = list(words)
        random.shuffle(words)

    created_files = []
    errors = []

    doc_q = docx.Document()
    apply_doc_style(doc_q, font_name, font_size)
    add_header_banner(doc_q, title_text, subtitle_text, logo_path, font_name)

    info_p = doc_q.add_paragraph()
    info_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    info_p.paragraph_format.space_before = Pt(4)
    info_p.paragraph_format.space_after = Pt(8)
    date_str = datetime.datetime.now().strftime("%Y. %m. %d")
    run = info_p.add_run(f"Name: ________________    Score: ______ / {len(words)}    Date: {date_str}")
    run.font.size = Pt(10)
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)

    _build_vocab_table(doc_q, words, exam_type,
                       show_answer=False,
                       show_first_letter=show_first_letter,
                       font_name=font_name)

    q_docx, q_pdf = build_paths(target_dir, base_name)
    ok, err = safe_save_docx(doc_q, q_docx)
    if not ok:
        return [], err
    created_files.append(q_docx)

    ans_base = f"{base_name}_정답"
    doc_a = docx.Document()
    apply_doc_style(doc_a, font_name, font_size)
    add_header_banner(doc_a, f"{title_text} [정답지]", subtitle_text, logo_path, font_name)

    _build_vocab_table(doc_a, words, exam_type,
                       show_answer=True,
                       font_name=font_name)

    a_docx, a_pdf = build_paths(target_dir, ans_base)
    ok, err = safe_save_docx(doc_a, a_docx)
    if not ok:
        errors.append(f"정답지 저장 실패: {err}")
    else:
        created_files.append(a_docx)

    if "PDF" in output_format:
        for dx, px in [(q_docx, q_pdf), (a_docx, a_pdf)]:
            if os.path.exists(dx):
                ok, msg = convert_to_pdf_safe(dx, px)
                if ok:
                    created_files.append(px)
                else:
                    errors.append(f"PDF 변환 실패: {msg}")

    err_msg = "\n".join(errors) if errors else ""
    return created_files, err_msg


# ═══════════════════════════════════════════
# v7.0: 예문 시험지 생성
# ═══════════════════════════════════════════

def create_sentence_test(target_dir, sentence_items, cfg,
                         test_mode="translation",
                         student_name="",
                         show_hint=False,
                         shuffle=False,
                         custom_title="",
                         custom_filename="",
                         logo_path="",
                         output_format="Word + PDF",
                         unit_names=None):
    """
    예문 시험지 생성 메인 함수 (v7.0)

    test_mode: "translation" (문장해석) | "cloze" (빈칸영작)
    """
    if not sentence_items:
        return [], "출제할 예문이 없습니다."

    os.makedirs(target_dir, exist_ok=True)

    font_name = cfg.get("font_name", "맑은 고딕")
    font_size = cfg.get("font_size", "10")
    if not logo_path:
        logo_path = cfg.get("last_logo", "")

    ts = datetime.datetime.now().strftime("%m%d_%H%M")
    mode_label = "문장해석" if test_mode == "translation" else "빈칸영작"

    if custom_title:
        title_text = custom_title
    else:
        title_text = f"Sentence Test — {mode_label}"
        if student_name:
            title_text += f" [{student_name}]"

    if custom_filename:
        base_name = sanitize_filename(custom_filename)
    else:
        name_part = f"_{student_name}" if student_name else ""
        base_name = f"SentTest_{mode_label}{name_part}_{ts}"

    subtitle_parts = []
    if unit_names:
        if len(unit_names) <= 3:
            subtitle_parts.append(" / ".join(unit_names))
        else:
            subtitle_parts.append(f"{unit_names[0]} ~ {unit_names[-1]} ({len(unit_names)}개 단원)")
    subtitle_parts.append(f"{mode_label} | {len(sentence_items)}문항")
    subtitle_text = "  |  ".join(subtitle_parts)

    if shuffle:
        sentence_items = list(sentence_items)
        random.shuffle(sentence_items)

    created_files = []
    errors = []

    doc_q = docx.Document()
    apply_doc_style(doc_q, font_name, font_size)
    add_header_banner(doc_q, title_text, subtitle_text, logo_path, font_name)

    info_p = doc_q.add_paragraph()
    info_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    info_p.paragraph_format.space_before = Pt(4)
    info_p.paragraph_format.space_after = Pt(8)
    date_str = datetime.datetime.now().strftime("%Y. %m. %d")
    run = info_p.add_run(f"Name: ________________    Score: ______ / {len(sentence_items)}    Date: {date_str}")
    run.font.size = Pt(10)
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)

    _build_sentence_table(doc_q, sentence_items, test_mode,
                          show_answer=False, show_hint=show_hint,
                          font_name=font_name)

    q_docx, q_pdf = build_paths(target_dir, base_name)
    ok, err = safe_save_docx(doc_q, q_docx)
    if not ok:
        return [], err
    created_files.append(q_docx)

    ans_base = f"{base_name}_정답"
    doc_a = docx.Document()
    apply_doc_style(doc_a, font_name, font_size)
    add_header_banner(doc_a, f"{title_text} [정답지]", subtitle_text, logo_path, font_name)

    _build_sentence_table(doc_a, sentence_items, test_mode,
                          show_answer=True, font_name=font_name)

    a_docx, a_pdf = build_paths(target_dir, ans_base)
    ok, err = safe_save_docx(doc_a, a_docx)
    if not ok:
        errors.append(f"정답지 저장 실패: {err}")
    else:
        created_files.append(a_docx)

    if "PDF" in output_format:
        for dx, px in [(q_docx, q_pdf), (a_docx, a_pdf)]:
            if os.path.exists(dx):
                ok, msg = convert_to_pdf_safe(dx, px)
                if ok:
                    created_files.append(px)
                else:
                    errors.append(f"PDF 변환 실패: {msg}")

    err_msg = "\n".join(errors) if errors else ""
    return created_files, err_msg
