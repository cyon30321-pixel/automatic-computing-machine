"""
강화된 텍스트 파싱 엔진 v5.5
— ①②③④⑤ / (A)(B)(C)(D)(E) 선지 자동 인식
— <조건>, [우리말], [보기], [정답] 블록 분리
— v5.4: 은/는 조사 정규식 수정, 추가 발문 패턴 보강
— v5.5: 선택지 마커 정규화 지원
"""

import re

CIRCLED_NUMS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def _is_valid_choice_block(text, marker_type):
    if marker_type == 'num':
        found = []
        for i, ch in enumerate(CIRCLED_NUMS[:5]):
            if ch in text:
                found.append(i + 1)
        for i in range(1, 6):
            if re.search(rf'(?:^|\s){i}\s*\)', text, re.MULTILINE):
                if i not in found:
                    found.append(i)
        found.sort()
        if len(found) < 3:
            return False
        consecutive = 1
        for j in range(1, len(found)):
            if found[j] == found[j-1] + 1:
                consecutive += 1
            else:
                consecutive = 1
            if consecutive >= 3:
                return True
        return consecutive >= 3

    elif marker_type == 'alpha':
        count = sum(1 for m in ["(A)", "(B)", "(C)", "(D)", "(E)"] if m in text)
        return count >= 3

    return False


def _extract_choices(choices_str, c_type):
    choices = ["", "", "", "", ""]

    if c_type == 'num':
        markers = []
        for i in range(5):
            circled = CIRCLED_NUMS[i]
            positions = []
            for m in re.finditer(re.escape(circled), choices_str):
                positions.append((m.start(), circled))
            for m in re.finditer(rf'(?:^|\s)({i+1}\s*\))', choices_str, re.MULTILINE):
                positions.append((m.start(), m.group(1)))
            if positions:
                positions.sort(key=lambda x: x[0])
                markers.append((i, positions[0][0], positions[0][1]))
        markers.sort(key=lambda x: x[1])

        for idx_in_markers, (choice_idx, start_pos, marker_text) in enumerate(markers):
            content_start = start_pos + len(marker_text)
            while content_start < len(choices_str) and choices_str[content_start] in ' \t':
                content_start += 1
            if idx_in_markers + 1 < len(markers):
                content_end = markers[idx_in_markers + 1][1]
            else:
                content_end = len(choices_str)
            val = choices_str[content_start:content_end].replace('\n', ' ').strip()
            if choice_idx < 5:
                choices[choice_idx] = val

    elif c_type == 'alpha':
        alpha_markers = ["(A)", "(B)", "(C)", "(D)", "(E)"]
        positions = []
        for i, marker in enumerate(alpha_markers):
            idx = choices_str.find(marker)
            if idx != -1:
                positions.append((i, idx, marker))
        positions.sort(key=lambda x: x[1])

        for idx_in_list, (choice_idx, start_pos, marker) in enumerate(positions):
            content_start = start_pos + len(marker)
            if idx_in_list + 1 < len(positions):
                content_end = positions[idx_in_list + 1][1]
            else:
                content_end = len(choices_str)
            val = choices_str[content_start:content_end].replace('\n', ' ').strip()
            if choice_idx < 5:
                choices[choice_idx] = val

    return choices


def parse_exam_text(raw_text):
    exam_data = []

    raw_text = re.sub(r'```[a-zA-Z]*\n?', '', raw_text)
    raw_text = raw_text.replace('```', '').replace("**", "")
    raw_text = re.sub(r"\[Section (Start|End)\]", "", raw_text, flags=re.IGNORECASE)
    raw_text = re.sub(r"^\s*[-=─]{5,}\s*$", "", raw_text, flags=re.MULTILINE)
    raw_text = re.sub(r"[-_]{3,}", " _______ ", raw_text)

    if re.search(r'\[\s*(?:Questions?\s*)?\d+\s*(?:-\s*\d+)?\s*\]', raw_text, flags=re.IGNORECASE):
        blocks = re.split(r'(?=\[\s*(?:Questions?\s*)?\d+\s*(?:-\s*\d+)?\s*\])', raw_text, flags=re.IGNORECASE)
    else:
        blocks = re.split(r"(?=다음\s*글을\s*읽고\s*물음에\s*답하시오)", raw_text)
        if len(blocks) == 1:
            blocks = [raw_text]

    global_q_num = 1

    for block in blocks:
        block = block.strip()
        if not block:
            continue

        passage_lines = []
        questions_raw = []
        current_q_lines = []
        in_question = False
        in_condition = False

        for line in block.split('\n'):
            line_str = line.strip()
            if not line_str:
                continue

            if re.search(r'(<조건>|\[조건\]|조건\s*:|\*\s*조건)', line_str):
                in_condition = True
            if re.search(r'(\[정답\]\s*:|정답\s*:|^①|^\(A\))', line_str):
                in_condition = False

            m = re.match(r'^(?:Q|문|문항)?\s*0*(\d+)[\.\\)]\s*(.*)', line_str, re.IGNORECASE)
            is_new_question = False

            if m:
                num = int(m.group(1))
                if not in_question or not in_condition:
                    is_new_question = True
                else:
                    if num > 9 or re.search(r'(하시오|고르시오|쓰시오|것[은는]?\.|인가\?|[은는]\?|적절한|옳[은는]|않[은는])$', line_str):
                        is_new_question = True
                        in_condition = False

            if is_new_question:
                if current_q_lines:
                    questions_raw.append("\n".join(current_q_lines))
                current_q_lines = [line_str]
                in_question = True
            else:
                if in_question:
                    current_q_lines.append(line_str)
                else:
                    passage_lines.append(line_str)

        if current_q_lines:
            questions_raw.append("\n".join(current_q_lines))

        passage_text = "\n".join(passage_lines).strip()
        passage_text = re.sub(
            r"^(?:\[.*?\]\s*)?(?:다음\s*글을\s*읽고\s*물음에\s*답하시오\.?)?",
            "", passage_text
        ).strip()

        questions = []
        for q_text_raw in questions_raw:
            m = re.match(r"^(?:Q|문|문항)?\s*0*(\d+)[\.\\)]\s*(.*)", q_text_raw.strip(), re.DOTALL | re.IGNORECASE)
            if not m:
                continue

            q_body = m.group(2)
            assigned_q_num = str(global_q_num)
            global_q_num += 1

            c_type = None
            body_part = q_body
            choices = []

            candidates = []
            for match_obj in re.finditer(r'①', q_body):
                rest = q_body[match_obj.start():]
                if _is_valid_choice_block(rest, 'num'):
                    candidates.append((match_obj.start(), 'num'))
            for match_obj in re.finditer(r'(?:^|\s)1\s*\)', q_body, re.MULTILINE):
                rest = q_body[match_obj.start():]
                if _is_valid_choice_block(rest, 'num'):
                    candidates.append((match_obj.start(), 'num'))
            for match_obj in re.finditer(r'\(A\)', q_body):
                rest = q_body[match_obj.start():]
                if _is_valid_choice_block(rest, 'alpha'):
                    candidates.append((match_obj.start(), 'alpha'))

            if candidates:
                candidates.sort(key=lambda x: x[0], reverse=True)
                best_start, best_type = candidates[0]
                body_part = q_body[:best_start].strip()
                choices_str = q_body[best_start:]
                c_type = best_type
                choices = _extract_choices(choices_str, c_type)

            questions.append({
                "num": assigned_q_num,
                "text": body_part,
                "choices": choices,
                "c_type": c_type
            })

        exam_data.append({"passage": passage_text, "questions": questions})

    return exam_data


def parse_question_with_choices(text):
    lines = text.strip().split("\n")
    question_lines = []
    choices = []
    choice_pattern = re.compile(r"^\s*[①②③④⑤]\s*")
    alpha_pattern = re.compile(r"^\s*[A-E][.)]\s*")
    for line in lines:
        if choice_pattern.match(line) or alpha_pattern.match(line):
            choices.append(line.strip())
        else:
            question_lines.append(line)
    return {"text": "\n".join(question_lines).strip(), "choices": choices}
