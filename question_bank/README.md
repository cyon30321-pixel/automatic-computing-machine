# 📚 영어 통합 문제은행 시스템 PRO v3.0

## 파일 구성
| 파일 | 설명 |
|------|------|
| `question_bank_pro.py` | 통합 문제은행 시스템 메인 프로그램 |
| `question_bank_pro_v5.db` | SQLite 데이터베이스 (문제 데이터) |

## 주요 기능
- 📝 지문+문제 자동 분해 및 DB 저장
- 🔍 고급 필터 검색 (학교급·유형·년도)
- 🛒 장바구니 기반 시험지/워크북 생성
- 📊 AI 오답 분석 및 시각화 (차트)
- 💾 DB 백업/복원 기능
- ⌨️ 키보드 단축키 (Ctrl+F 검색, Ctrl+S 생성)

## 설치 및 실행
```bash
pip install python-docx docx2pdf matplotlib
python question_bank_pro.py
```