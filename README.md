# Inbound Tradebook (China TRS 관리 웹앱)

Excel 기반으로 관리하던 China 주식 TRS(Total Return Swap) Inbound Tradebook을
SQLite DB + Streamlit 웹앱으로 전환한 프로젝트입니다.

- 신규거래(OPEN) 입력/조회/수정
- UNWIND(청산) 입력 및 최종본 A~Z 컬럼 자동 계산
- "다가오는 Reset" 알림 (D-day 표시, 지난 리셋 경고)
- 블룸버그 의존 항목은 `_xll.BDH` 수식이 포함된 엑셀로 내보내 터미널 PC에서 계산
- 휴일 캘린더 기반 WORKDAY 계산

> 이 저장소의 데이터는 모두 가상 샘플입니다 (계좌번호·펀드명·티커 전부 가상).

## 실행 방법

```bash
pip install -r requirements.txt
python seed_sample_data.py   # 가상 샘플 데이터 채우기 (최초 1회)
streamlit run app.py
```
