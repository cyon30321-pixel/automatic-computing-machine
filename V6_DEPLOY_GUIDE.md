# v6.0 배포 가이드 및 변경 로그

## 📋 변경 파일 목록 (10개)

| # | 파일 | 상태 | 설명 |
|---|------|------|------|
| 1 | `exam_bank/__init__.py` | 수정 | 버전 6.0.0 |
| 2 | `exam_bank/constants.py` | 수정 | APP_NAME v6.0, VOCAB_CATEGORIES 추가 |
| 3 | `exam_bank/app.py` | 수정 | 8번째 탭(📖 단어장) 등록, 상태바 단어 통계 |
| 4 | `exam_bank/models/database.py` | 수정 | 단어장 5개 테이블, Soft Delete(is_active), 마이그레이션 |
| 5 | `exam_bank/models/vocabulary.py` | **신규** | 단어장 CRUD, bulk import, 학습 통계, 오답 추출 |
| 6 | `exam_bank/services/docx_utils.py` | **신규** | 공통 Word 유틸리티 (배너, 스타일, PDF, sanitization) |
| 7 | `exam_bank/services/generator.py` | 수정 | custom_title/custom_filename, sanitization, PermissionError 방어 |
| 8 | `exam_bank/services/vocab_generator.py` | **신규** | 단어 시험지 생성 (2단 레이아웃, 영↔한, 정답지) |
| 9 | `exam_bank/views/bank_tab.py` | 수정 | 드래그 선택, Ctrl+A, 제목/파일명 지정, 배너, LabelFrame |
| 10 | `exam_bank/views/vocab_tab.py` | **신규** | 단어장 관리 탭 (AI 추출, Day 조합, 시험 생성) |

## 🚀 v6.0 신규 기능 (5개)

### Feature 1: 문제관리 트리뷰 드래그 다중 선택
- `<B1-Motion>` 이벤트로 드래그 연속 선택
- 디바운싱: 같은 행 반복 처리 방지 (`_drag_last_row`)
- `Ctrl+A` 전체 선택 (포커스 체크로 Entry 충돌 방지)

### Feature 2: 더블클릭 장바구니 담기 피드백 개선
- 기존 기능 유지 + 상태바에 "✓ N개 항목 장바구니에 추가됨" 표시

### Feature 3: 시험지 제목/파일명 커스터마이징
- `⚙️ 출력 고급 설정` LabelFrame 내 Entry 2개
- 파일명 sanitization: `re.sub(r'[\\/:*?"<>|]', '_', filename)`
- PermissionError 방어: try-except로 "파일을 먼저 닫아주세요" 안내

### Feature 4: 문제관리 배너 삽입
- `⚙️ 출력 고급 설정` 내 배너 선택/제거 버튼
- 기존 `add_header_banner()` 재사용

### Feature 5: 단어장 시스템 (📖 단어장 탭)
- **단어장 관리**: Book → Unit(Day) → Word 3단 계층
- **AI 추출 가져오기**: Claude 프롬프트 → JSON → 미리보기 → DB 저장
- **Day 조합 출제**: 여러 Day를 조합하여 출제 범위 설정
- **시험지 생성**: 2단 FactoryVoca 스타일, 영→한/한→영, 정답지 자동 생성
- **오답 재시험**: 학생별 틀린 단어만 추출하여 재시험지 생성
- **파일 가져오기**: Excel/CSV에서 일괄 등록

## 🏗️ 아키텍처 개선

### DB 스키마 v3
- `vocab_books`: 단어장 메타 (Soft Delete)
- `vocab_units`: 단원/Day (Soft Delete)
- `vocab_words`: 개별 단어
- `vocab_exam_records`: 시험 기록
- `vocab_exam_items`: 개별 오답 추적

### 코드 분리
- `docx_utils.py`: generator.py + vocab_generator.py 공통 함수 분리
- `vocabulary.py`: 단어장 전용 모델 레이어

## 📦 배포 방법

1. v6_output 내 10개 파일을 기존 프로젝트에 덮어쓰기
2. 기존 DB는 자동 마이그레이션 (ALTER TABLE + CREATE TABLE IF NOT EXISTS)
3. 추가 패키지 불필요 (기존 python-docx, docx2pdf 그대로)

## ⚠️ 호환성 참고
- v5.5 DB → v6.0 자동 업그레이드 (마이그레이션 함수 내장)
- 기존 기능 100% 하위 호환
- 단어장 기능은 독립 테이블이므로 기존 데이터에 영향 없음
