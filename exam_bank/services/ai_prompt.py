"""
AI 프롬프트 자동 생성 서비스 v5.0

목적: AI(ChatGPT/Claude)에게 영어 변형문제를 요청할 때,
파이썬 파싱 엔진이 깔끔하게 인식할 수 있는 포맷으로
문제를 출력하도록 하는 프롬프트를 자동 생성.

사용 흐름:
1. 사용자가 원본 지문을 붙여넣음
2. 문제 유형/개수/난이도 선택
3. 프롬프트 자동 생성 → 복사 → AI에 붙여넣기
4. AI 결과를 시험지 생성 탭에 붙여넣으면 깔끔 출력
"""


# ─── 문제 유형 정의 ───
QUESTION_TYPES = {
    "빈칸 추론": {
        "label": "빈칸 추론 (Blank Inference)",
        "desc": "지문의 핵심 표현을 빈칸으로 만들고 5지선다로 출제",
        "instruction": "지문에서 핵심 어구나 문장을 빈칸(_______)으로 만들고, 빈칸에 들어갈 가장 적절한 것을 고르는 문제를 만들어라.",
    },
    "주제/요지": {
        "label": "주제/요지 파악",
        "desc": "글의 주제, 요지, 제목을 묻는 문제",
        "instruction": "글의 주제(topic), 요지(main idea), 또는 적절한 제목(title)을 묻는 문제를 만들어라.",
    },
    "어법 (문법)": {
        "label": "어법 정확성 판단",
        "desc": "밑줄 친 부분의 어법 정확성을 판단",
        "instruction": "지문의 일부 표현에 밑줄을 긋고, 어법상 옳은 것/틀린 것을 고르는 문제를 만들어라. (A)/(B)/(C) 형식으로 각각 2개 선택지를 주고, 조합을 ①②③④⑤로 출제.",
    },
    "어휘": {
        "label": "어휘 적절성 판단",
        "desc": "밑줄 친 어휘의 적절성 판단",
        "instruction": "지문의 핵심 어휘에 밑줄을 긋고, 문맥상 적절한 어휘/적절하지 않은 어휘를 고르는 문제를 만들어라.",
    },
    "순서 배열": {
        "label": "문장 순서 배열",
        "desc": "주어진 문장 다음에 올 순서를 배열",
        "instruction": "지문을 (A), (B), (C) 세 단락으로 나누고, 주어진 글 다음에 이어질 순서로 가장 적절한 것을 고르는 문제를 만들어라.",
    },
    "문장 삽입": {
        "label": "문장 삽입 위치",
        "desc": "주어진 문장이 들어가기에 가장 적절한 곳",
        "instruction": "지문에 ①②③④⑤ 위치를 표시하고, 주어진 문장이 들어가기에 가장 적절한 곳을 고르는 문제를 만들어라.",
    },
    "내용 일치/불일치": {
        "label": "내용 일치/불일치",
        "desc": "글의 내용과 일치하는/일치하지 않는 것 고르기",
        "instruction": "글의 내용과 일치하는 것 또는 일치하지 않는 것을 고르는 문제를 만들어라.",
    },
    "서술형": {
        "label": "서술형 문제",
        "desc": "영작문, 요약, 빈칸 서술 등",
        "instruction": "서술형 문제를 만들어라. (조건)을 명시하고, 학생이 직접 영어로 답을 작성하도록 출제.",
    },
    "종합 (혼합)": {
        "label": "종합 혼합 출제",
        "desc": "다양한 유형을 섞어서 출제",
        "instruction": "빈칸 추론, 어법, 어휘, 내용 일치, 순서 배열, 문장 삽입 등 다양한 유형을 골고루 섞어서 출제하라.",
    },
}


def generate_ai_prompt(passage_text, question_types, num_questions=10,
                       difficulty="중", school_level="고등학교",
                       grade="2학년", extra_instructions=""):
    """
    AI에게 보낼 프롬프트를 자동 생성합니다.

    Args:
        passage_text: 원본 영어 지문
        question_types: 선택된 문제 유형 리스트 (예: ["빈칸 추론", "어법 (문법)"])
        num_questions: 문제 수
        difficulty: 난이도 ("하", "중", "상", "최상")
        school_level: 학교급 ("중학교", "고등학교")
        grade: 학년
        extra_instructions: 추가 지시사항

    Returns:
        str: 완성된 AI 프롬프트
    """

    # 문제 유형별 지시사항 조합
    type_instructions = []
    for qt in question_types:
        info = QUESTION_TYPES.get(qt, {})
        if info:
            type_instructions.append(f"- {info['label']}: {info['instruction']}")

    types_block = "\n".join(type_instructions) if type_instructions else "- 다양한 유형을 골고루 섞어서 출제하라."

    prompt = f"""============================================================
🎯 역할 (ROLE)
============================================================
너는 한국의 {school_level} {grade} 영어 내신 시험을 전문적으로 출제하는
20년 경력의 영어 교육 전문가이다.
주어진 영어 지문을 바탕으로 변형 문제를 만들어라.

============================================================
📋 출력 형식 규칙 (★★★ 반드시 지켜야 함 ★★★)
============================================================
이 문제는 파이썬 프로그램으로 자동 파싱되어 시험지로 출력된다.
따라서 아래 형식을 정확하게 지켜야 한다. 형식이 틀리면 시험지가 깨진다.

[규칙 1] 지문은 맨 위에 한 번만 적는다.
  "다음 글을 읽고 물음에 답하시오."로 시작한다.

[규칙 2] 각 문제는 반드시 "번호." 으로 시작한다.
  예: 1. / 2. / 3.  (Q1, 문1, (1) 등 사용 금지)

[규칙 3] 객관식 선택지는 반드시 원형 숫자 ①②③④⑤ 를 사용한다.
  각 선지는 새 줄에 적는다.
  예:
  ① The boy was running fast.
  ② She decided to leave early.
  ③ They have been studying hard.
  ④ He will not come tomorrow.
  ⑤ We should consider the option.

[규칙 4] (A)(B)(C) 형식 문제(어법/어휘 등)도 선택지는 ①②③④⑤로 적는다.
  예:
  (A), (B), (C)의 각 괄호 안에서 어법에 맞는 표현을 고른 것은?
  ① A - B - C
  ② A - B - D
  ...

[규칙 5] <조건> 블록이 필요한 경우 아래처럼 적는다.
  <조건>
  - 주어진 단어를 반드시 사용할 것
  - 3문장 이내로 쓸 것

[규칙 6] 서술형 문제는 선택지 없이 문제만 적는다.

[규칙 7] 정답 및 해설은 모든 문제 뒤에 별도 섹션으로 적는다.
  형식: "번호. 정답: ① (또는 ②③④⑤)" + 해설
  예:
  1. 정답: ③
  해설: 빈칸에는 '역설적으로'라는 의미의 paradoxically가 적절하다.

[규칙 8] 마크다운(**, ```, --- 등) 절대 사용 금지.
         구분선(===, ---) 사용 금지. 순수 텍스트만 사용.

============================================================
📝 출제 조건
============================================================
- 문제 수: {num_questions}문항
- 난이도: {difficulty}
- 대상: {school_level} {grade}
- 문제 유형:
{types_block}

============================================================
📖 출력 예시 (이 형식을 정확히 따라라)
============================================================
다음 글을 읽고 물음에 답하시오.

[여기에 지문이 들어감]

1. 윗글의 빈칸에 들어갈 말로 가장 적절한 것은?
① increased dramatically
② remained unchanged
③ shifted paradoxically
④ decreased slightly
⑤ expanded rapidly

2. 윗글의 밑줄 친 (A), (B), (C)에서 어법에 맞는 표현으로 가장 적절한 것은?
① which - adapting - their
② that - adapted - its
③ which - adapted - their
④ that - adapting - its
⑤ which - adapting - its

3. 윗글의 내용과 일치하지 않는 것은?
① The study was conducted over five years.
② Participants showed improvement in memory.
③ The control group received no treatment.
④ Results were published in a medical journal.
⑤ All subjects were over the age of 60.

[정답 및 해설]
1. 정답: ③
해설: 문맥상 '역설적으로 변화했다'는 의미가 적절하다.

2. 정답: ③
해설: (A) 선행사가 사물이므로 which, (B) 과거 사실이므로 adapted, (C) 복수 주어이므로 their

3. 정답: ④
해설: 지문에서 결과가 학술지에 발표되었다는 내용은 언급되지 않았다.

============================================================
🔤 원본 지문 (이 지문을 바탕으로 변형문제를 만들어라)
============================================================
{passage_text.strip()}
"""

    if extra_instructions.strip():
        prompt += f"""

============================================================
💬 추가 지시사항
============================================================
{extra_instructions.strip()}
"""

    return prompt.strip()


def get_question_type_list():
    """UI에서 사용할 문제 유형 목록 반환"""
    return list(QUESTION_TYPES.keys())


def get_question_type_info(type_name):
    """특정 문제 유형의 상세 정보 반환"""
    return QUESTION_TYPES.get(type_name, {})
