# rehab-collector

회생절차 진행 기업 정보를 공개 데이터에서 자동 수집하는 도구 모음입니다.

두 갈래로 나뉩니다.

| | 대상 | 데이터 경로 | 상태 |
|---|---|---|---|
| **트랙 A** | 상장사·공시대상 법인 | DART OpenAPI (공식 API) | 코드 완성 |
| **트랙 B** | 그 외 (비상장 중소기업) | 대법원 M&A 매각공고 | 코드 완성 |

두 트랙 모두 **같은 16열 엑셀**을 내보내므로 결과를 그대로 이어 붙일 수 있습니다.

---

## 빠른 시작

### 설치 (한 번만)

```bash
git clone https://github.com/dono011/rehab-collector.git
cd rehab-collector
pip3 install -r track_b_court/requirements.txt --break-system-packages
```

### 트랙 B — 법원 M&A 매각공고 (API 키 불필요)

```bash
cd track_b_court
python3 collect_court.py --check      # 접근 확인
python3 collect_court.py --pages 5    # 수집
python3 collect_court.py --codes      # 업종·법원 코드표
```

결과: `output/법원공고_회생업체_YYYYMMDD.xlsx`

### 트랙 A — DART 공시 (API 키 필요)

키 발급(무료): https://opendart.fss.or.kr → 인증키 신청

```bash
cd track_a_dart
export DART_API_KEY="발급받은40자리키"
python3 collect_rehab.py --test       # 연결 확인
python3 collect_rehab.py --days 90    # 수집
```

결과: `output/회생신청_정보_YYYYMMDD.xlsx`

---

## 각 트랙 상세

- [`track_a_dart/README.md`](track_a_dart/README.md) — DART 수집기
- [`track_b_court/README.md`](track_b_court/README.md) — 법원 공고 수집기

---

## 수집 항목 (16열)

| 항목 | 트랙 A | 트랙 B |
|---|---|---|
| 회사명 / 관할법원 / 사건상태 | ✅ | ✅ |
| 사업자번호 / 대표자명 / 전화번호 / 소재지 | ✅ | ❌ |
| 자본금 / 채무액 | ✅ | ❌ |
| 업종 | DART 업종코드 | 법원 분류 + 회사명 |
| 사건번호 | 공시 원문 링크 | 공고 원문 링크 |

트랙 B에서 비는 항목은 등기부·협회 조회로만 채울 수 있습니다.

---

## 값을 그대로 믿으면 안 되는 항목

- **채무액** — 재무제표 부채총계입니다. 회생채권 확정액이 아니고 **우발부채(PF 지급보증 등)가 빠져 있습니다.** 실제 규모는 훨씬 클 수 있습니다.
- **업종·관련도** — 자동 분류·추정입니다. 특정 업종 면허의 실제 보유 여부는 소관 협회 조회로 확인해야 합니다.
- **법원 업종 분류가 느슨합니다** — 실제 목록을 보면 상당수가 "기타"로 등록되어 있어, 업종으로 좁혀 받으면 놓칩니다. 전체를 받아 후처리하는 편이 낫습니다.

---

## 법적 준수

- **트랙 A**는 금융감독원이 제공하는 **공식 OpenAPI**를 사용합니다.
- **트랙 B**는 매 실행 시 `robots.txt`를 확인하고, 금지된 경로면 해당 사이트를 건너뜁니다. 요청 간격은 1.2초입니다.
  - 대상 경로(`/portal/notice/...`)는 `urllib.robotparser` 판정 결과 **허용**입니다.
- 공개된 공고만 수집하며, 민간 집계 사이트는 **DB제작자 권리** 문제로 대상에서 제외했습니다.

---

## 요구사항

```
Python 3.8+
requests / openpyxl          (트랙 A)
requests / openpyxl / bs4    (트랙 B)
```

---

## 매일 자동 실행

```bash
crontab -e
```

```
0  2 * * * cd ~/rehab-collector/track_a_dart && DART_API_KEY="키" /usr/bin/python3 collect_rehab.py --days 7 >> output/cron.log 2>&1
30 2 * * * cd ~/rehab-collector/track_b_court && /usr/bin/python3 collect_court.py --pages 2 >> output/cron.log 2>&1
```

---

## 업데이트 받기

```bash
cd ~/rehab-collector
git pull
```
