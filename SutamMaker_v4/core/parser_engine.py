"""
SutamMaker v4 — core/parser_engine.py
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Text parsing engine: raw exam text -> structured question list.
No Tkinter / docx / matplotlib dependencies.
"""

import re

from SutamMaker_v4.config import CIRCLE_NUMS


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# parse_exam_text
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def parse_exam_text(raw_text):
    """Parse raw text into a structured list of question dicts."""
    raw_text = raw_text.lstrip("\ufeff")
    raw_text = raw_text.replace("\r\n", "\n").replace("\r", "\n")
    raw_text = raw_text.replace("\u00a0", " ")
    raw_text = raw_text.replace("\u3000", " ")
    raw_text = re.sub(r"```[a-zA-Z]*\n?", "", raw_text)
    raw_text = raw_text.replace("```", "").replace("**", "").replace("$", "")

    SPLIT_PAT = r"\n(?=(?:Q|문|문항)?\s{0,3}\d{1,3}[\.\):\s]\s*[가-힣A-Za-z<\(\[（「【])"
    MATCH_PAT = r"^\s*(?:Q|문|문항)?\s{0,3}([0-9]{1,3})[\.\):]?\s+(.*)"

    splits = re.split(SPLIT_PAT, "\n" + raw_text)
    result = []

    for chunk in splits:
        if not chunk.strip():
            continue
        m = re.match(MATCH_PAT, chunk, re.DOTALL | re.IGNORECASE)
        if not m:
            continue

        q_num = m.group(1)
        body = m.group(2)

        # ── image tag ──
        image_path = None
        img_match = re.search(r"\[IMAGE\s*:\s*(.+?)\]", body, re.IGNORECASE)
        if img_match:
            image_path = img_match.group(1).strip()
            body = body.replace(img_match.group(0), "").strip()

        # ── choices ──
        choices = []
        c_start = re.search(r"(①|1\s*\))", body)
        if c_start:
            choices_str = body[c_start.start():]
            body = body[:c_start.start()].strip()
            for i in range(1, 6):
                circle = CIRCLE_NUMS[i - 1]
                if i < 5:
                    next_circle = CIRCLE_NUMS[i]
                    pat = f"({re.escape(circle)}|{i}\\s*\\)).*?(?={re.escape(next_circle)}|{i+1}\\s*\\)|$)"
                else:
                    pat = f"({re.escape(circle)}|{i}\\s*\\)).*"
                cm = re.search(pat, choices_str, re.DOTALL)
                if cm:
                    val = re.sub(f"^({re.escape(circle)}|{i}\\s*\\))\\s*", "", cm.group(0)).replace("\n", " ").strip()
                    choices.append(val)
                else:
                    choices.append("")

        # ── graph tags ──
        lines = body.split("\n")
        clean_lines = []
        graph_parts = []
        in_graph = False
        for line in lines:
            up = line.upper()
            if re.search(r"\[\s*(?:DUAL_)?(?:GRAPH|PPF|ASAD|AS.AD|LORENZ|PHILLIPS|TAX|TARIFF|EXTERNALITY|INDIFFERENCE)", up):
                in_graph = True
                graph_parts.append(line)
                if line.count("[") == line.count("]") and "]" in line:
                    in_graph = False
            elif in_graph:
                graph_parts.append(line)
                if "]" in line:
                    in_graph = False
            elif re.search(
                r"^\s*[\|\[\$]?\s*(?:Qd|Qs|P|Shift_[SD]|Price_|PPF_SHIFT|GINI|COMPARE_GINI|PHILLIPS_SHIFT|Shift_AD|Shift_SRAS)\s*=",
                line, re.IGNORECASE,
            ):
                graph_parts.append(line)
            else:
                clean_lines.append(line)
        graph_tag = " ".join(graph_parts) if graph_parts else None
        body = "\n".join(clean_lines).strip()

        # ── jesi (passage) ──
        jesi = ""
        jm = re.search(r"<제시문>\s*(.*?)\s*</제시문>", body, re.DOTALL | re.IGNORECASE)
        if jm:
            jesi = jm.group(1).strip()
            body = body.replace(jm.group(0), "").strip()

        # ── bogi (example box) ──
        bogi = ""
        bogi_tag = re.search(r"^\s*[\[<]\s*보\s*기\s*[\]>][ \t]*\n?", body, re.MULTILINE)
        if bogi_tag:
            bogi = body[bogi_tag.end():].strip()
            body = body[:bogi_tag.start()].strip()
        else:
            bogi_lines = re.findall(r"^[ㄱㄴㄷㄹㅁㅂ]\s*[\.\)].*", body, re.MULTILINE)
            if bogi_lines:
                bogi = "\n".join(bogi_lines)
                for bl in bogi_lines:
                    body = body.replace(bl, "")
                body = body.strip()

        body = re.sub(r"[\[<]\s*보\s*기\s*[\]>]", "보기", body).strip()

        # ── (가)(나)(다) blocks -> jesi ──
        if not jesi:
            gana_pat = r"^\s*[\(（]([가나다라마바사아자차카타파하])\s*[\)）]\s+.+"
            gana_lines = re.findall(gana_pat, body, re.MULTILINE)
            if len(gana_lines) >= 2:
                gana_full = re.findall(
                    r"^\s*[\(（][가나다라마바사아자차카타파하]\s*[\)）]\s+.+",
                    body, re.MULTILINE,
                )
                jesi = "\n".join(line.strip() for line in gana_full)
                for gl in gana_full:
                    body = body.replace(gl, "")
                body = body.strip()

        # ── tables ──
        tables = []
        final_lines = []
        current_table = []
        for line in body.split("\n"):
            ls = line.strip()
            if "|" in ls and not re.match(r"^[-:\s\|\+]+$", ls):
                current_table.append([c.strip() for c in ls.strip("|").split("|")])
            elif "|" in ls and re.match(r"^[-:\s\|\+]+$", ls):
                continue
            else:
                if current_table:
                    tables.append(current_table)
                    final_lines.append("[TABLE_PLACEHOLDER]")
                    current_table = []
                final_lines.append(line)
        if current_table:
            tables.append(current_table)
            final_lines.append("[TABLE_PLACEHOLDER]")

        body = re.sub(r"\n{3,}", "\n\n", "\n".join(final_lines).strip())

        result.append({
            "num": q_num,
            "text": body,
            "tables": tables,
            "bogi": bogi,
            "jesi": jesi,
            "graph_tag": graph_tag,
            "image_path": image_path,
            "choices": choices,
        })

    return result


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# parse_answer_by_question
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def parse_answer_by_question(raw_ans):
    """Split answer/explanation text per question number.
    Returns: {"1": "1번 정답: ③\\n[해설] ...", "2": "2번 정답: ⑤\\n...", ...}
    """
    if not raw_ans or not raw_ans.strip():
        return {}
    result = {}
    split_pat = r"\n(?=\s*\d{1,3}\s*(?:번|[\.\)])\s*)"
    chunks = re.split(split_pat, "\n" + raw_ans)
    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk:
            continue
        m = re.match(r"^\s*(\d{1,3})\s*(?:번|[\.\)])\s*(.*)", chunk, re.DOTALL)
        if m:
            q_num = m.group(1).lstrip("0") or "0"
            result[q_num] = chunk
    return result
