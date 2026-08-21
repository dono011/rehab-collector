#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
트랙 A — DART 기반 회생절차 공시 자동 수집기

금융감독원 전자공시(DART) OpenAPI로 회생 관련 공시를 수집해
엑셀로 저장한다. 크롤링이 아니라 공식 API를 쓰므로
robots.txt / IP 차단 / 법적 리스크가 없다.

사용법:
    export DART_API_KEY="발급받은키"
    python3 collect_rehab.py --test          # API 키 확인만
    python3 collect_rehab.py                 # 최근 30일
    python3 collect_rehab.py --days 90       # 최근 90일
    python3 collect_rehab.py --from 20260101 --to 20260819
    python3 collect_rehab.py --no-financials # 재무조회 생략(빠름)

API 키 발급(무료): https://opendart.fss.or.kr → 인증키 신청
"""

import argparse
import json
import os
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = "https://opendart.fss.or.kr/api"
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
import categories as CAT  # noqa: E402
OUT_DIR = HERE / "output"
SEEN_FILE = OUT_DIR / "seen.json"
CACHE_FILE = OUT_DIR / "company_cache.json"

# 공시 제목에서 찾을 키워드
KEYWORDS = ["회생", "파산", "워크아웃", "기업개선"]

# 검색할 공시 유형 (B=주요사항보고서, I=거래소공시)
PBLNTF_TYPES = ["B", "I"]

# 엑셀 컬럼 (원 기획서 16개 항목)
COLUMNS = [
    "수집일자", "회사명", "사업자번호", "대표자명", "관할법원", "사건번호",
    "신청일", "결정일", "사건상태", "면허/업종", "분류", "판단근거", "자본금", "채무액",
    "소재지", "전화번호", "데이터출처", "비고",
]

# 한국표준산업분류 접두어 → 전기 관련도
# 2026-08-21 확대: 전기만 보다가 사장님 지시로 건설·소방·설비통신까지 넣었다.
# 긴 접두어가 먼저 와야 한다 (423 이 42 보다 먼저 걸려야 전기로 잡힌다).
INDUTY_GROUPS = [
    ("4231", "전기공사업", "전기", 3),
    ("4232", "통신공사업", "설비통신", 3),
    ("423",  "전기·통신공사업", "전기", 3),
    ("422",  "건물설비 설치공사업(배관·냉난방·소방)", "설비통신", 3),
    ("421",  "기반조성·시설물축조 공사업", "건설", 3),
    ("424",  "실내건축·건축마무리 공사업", "건설", 3),
    ("425",  "건설 시설물 유지관리업", "건설", 3),
    ("42",   "전문직별 공사업", "건설", 3),
    ("41",   "종합 건설업", "건설", 3),
    ("351",  "전기업(발전·송배전)", "전기", 2),
    ("28",   "전기장비 제조업", "전기", 2),
    ("261",  "반도체·전자부품", "전기", 1),
]

DART_ERRORS = {
    "010": "등록되지 않은 키입니다",
    "011": "사용할 수 없는 키입니다 (오픈API 이용 등록 확인)",
    "012": "접근할 수 없는 IP입니다",
    "013": "조회된 데이터가 없습니다",
    "020": "요청 제한을 초과했습니다 (일 20,000건)",
    "100": "필드값이 부적절합니다",
    "800": "시스템 점검 중입니다",
    "900": "정의되지 않은 오류입니다",
    "901": "사용자 계정의 개인정보보유기간이 만료되었습니다",
}


# ─────────────────────────────────────────────────────────
# 기본 유틸
# ─────────────────────────────────────────────────────────

def log(msg=""):
    print(msg, flush=True)


def get_api_key(cli_key=None):
    """API 키를 인자 → 환경변수 → 파일 순으로 찾는다."""
    if cli_key:
        return cli_key.strip()
    env = os.environ.get("DART_API_KEY", "").strip()
    if env:
        return env
    keyfile = HERE / "dart_key.txt"
    if keyfile.exists():
        return keyfile.read_text(encoding="utf-8").strip()
    return None


def call_api(endpoint, params, key, retries=3):
    """DART API 호출. (data, error_message) 반환."""
    params = dict(params, crtfc_key=key)
    url = f"{BASE}/{endpoint}"

    for attempt in range(retries):
        try:
            r = requests.get(url, params=params, timeout=30)
            r.raise_for_status()
            data = r.json()
        except requests.RequestException as e:
            if attempt < retries - 1:
                wait = 2 ** (attempt + 1)
                log(f"    네트워크 오류, {wait}초 후 재시도… ({e})")
                time.sleep(wait)
                continue
            return None, f"네트워크 오류: {e}"
        except ValueError:
            return None, "응답이 JSON이 아닙니다 (점검 중일 수 있음)"

        status = data.get("status")
        if status == "000":
            return data, None
        if status == "013":                      # 데이터 없음은 정상
            return None, None
        return None, DART_ERRORS.get(status, data.get("message", f"오류 {status}"))

    return None, "재시도 횟수 초과"


def load_json(path, default):
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (ValueError, OSError):
            pass
    return default


def save_json(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


# ─────────────────────────────────────────────────────────
# 수집
# ─────────────────────────────────────────────────────────

def split_periods(bgn_de, end_de, max_days=80):
    """DART는 corp_code 없이 조회하면 **3개월까지만** 허용한다.
    (2026-08-21 확인: 6개월로 부르면 status 100 "검색기간은 3개월만 가능합니다")
    그래서 긴 기간은 잘라서 여러 번 부른다. 여유를 두고 80일씩 자른다."""
    from datetime import datetime, timedelta
    b = datetime.strptime(bgn_de, "%Y%m%d")
    e = datetime.strptime(end_de, "%Y%m%d")
    out = []
    while b <= e:
        chunk_end = min(b + timedelta(days=max_days - 1), e)
        out.append((b.strftime("%Y%m%d"), chunk_end.strftime("%Y%m%d")))
        b = chunk_end + timedelta(days=1)
    return out


def fetch_disclosures(key, bgn_de, end_de):
    """기간 내 공시 중 회생 관련 건만 추린다."""
    found = []
    seen_rcept = set()
    errors = []

    for ty in PBLNTF_TYPES:
      for sub_bgn, sub_end in split_periods(bgn_de, end_de):
        page = 1
        total_page = 1
        while page <= total_page:
            data, err = call_api("list.json", {
                "bgn_de": sub_bgn,
                "end_de": sub_end,
                "pblntf_ty": ty,
                "page_no": page,
                "page_count": 100,
            }, key)

            if err:
                # "조회된 데이터가 없습니다"는 오류가 아니라 정상이다
                if "없습니다" in str(err) and "기간" not in str(err):
                    break
                msg = f"[{ty}] {sub_bgn}~{sub_end} {page}페이지: {err}"
                log("  오류 " + msg)
                errors.append(msg)
                break
            if not data:
                break

            total_page = int(data.get("total_page", 1) or 1)
            for item in data.get("list", []):
                name = item.get("report_nm", "")
                if not any(kw in name for kw in KEYWORDS):
                    continue
                rcept = item.get("rcept_no")
                if rcept in seen_rcept:
                    continue
                seen_rcept.add(rcept)
                found.append(item)

            log(f"  [{ty}] {sub_bgn}~{sub_end} {page}/{total_page}페이지 — 누적 {len(found)}건")
            page += 1
            time.sleep(0.15)          # API 예의상 간격

    if errors:
        log("")
        log(f"  ⚠️ 조회 실패 {len(errors)}건 — 아래 결과는 일부만 담고 있습니다.")

    return found


def fetch_company(key, corp_code, cache):
    """기업개황 조회 (캐시 사용)."""
    if corp_code in cache:
        return cache[corp_code]
    data, err = call_api("company.json", {"corp_code": corp_code}, key)
    info = data if data else {"_error": err or "조회 불가"}
    cache[corp_code] = info
    time.sleep(0.15)
    return info


def fetch_financials(key, corp_code, year):
    """자본금 / 부채총계를 최근 사업보고서에서 가져온다."""
    for y in (year, year - 1):
        data, _ = call_api("fnlttSinglAcnt.json", {
            "corp_code": corp_code,
            "bsns_year": str(y),
            "reprt_code": "11011",           # 사업보고서
        }, key)
        time.sleep(0.15)
        if not data:
            continue

        result = {}
        for row in data.get("list", []):
            nm = (row.get("account_nm") or "").replace(" ", "")
            amount = row.get("thstrm_amount", "")
            if not amount:
                continue
            if nm == "자본금" and "자본금" not in result:
                result["자본금"] = amount
            elif nm == "부채총계" and "부채총계" not in result:
                result["부채총계"] = amount
        if result:
            result["기준연도"] = y
            return result
    return {}


# ─────────────────────────────────────────────────────────
# 가공
# ─────────────────────────────────────────────────────────

def classify_induty(code):
    """업종코드 → (설명, 분류, 관련도 0~3)"""
    code = (code or "").strip()
    if not code:
        return "미상", "", 0
    for prefix, label, cat, score in INDUTY_GROUPS:
        if code.startswith(prefix):
            return f"{label} ({code})", cat, score
    return f"기타 ({code})", "", 0


def decide_category(corp_name, induty_label, induty_cat, induty_score):
    """업종코드 분류와 회사명 분류를 합친다.

    DART 업종코드는 상당수가 "기타"로 뭉뚱그려져 있어(2026-08-21 확인)
    업종만 믿으면 놓친다. 회사명 쪽 판단(categories.py)을 같이 쓴다.
    """
    by_name = CAT.classify(corp_name)
    cats = list(by_name["카테고리"])
    score = by_name["점수"]
    if induty_cat:
        if induty_cat in cats:
            cats.remove(induty_cat)
        cats.insert(0, induty_cat)          # 업종코드 쪽을 앞에 둔다
        score = max(score, induty_score)
    reason = by_name["근거"]
    if induty_cat:
        reason = (f"업종코드:{induty_label}({induty_score})"
                  + (", " + reason if reason else ""))
    return "/".join(cats), score, reason


def guess_status(report_nm):
    """공시 제목에서 사건 진행 단계를 추정."""
    pairs = [
        ("개시신청", "신청"), ("개시결정", "개시"), ("개시 결정", "개시"),
        ("인가", "인가"), ("종결", "종결"), ("폐지", "폐지"),
        ("기각", "기각"), ("취하", "취하"), ("파산", "파산"),
    ]
    for needle, label in pairs:
        if needle in report_nm:
            return label
    return "확인필요"


def fmt_won(value):
    """숫자 문자열을 억원 단위로."""
    try:
        n = int(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return str(value or "")
    return f"{n / 100_000_000:,.1f}억" if abs(n) >= 100_000_000 else f"{n:,}원"


def fmt_date(s):
    s = (s or "").strip()
    return f"{s[:4]}-{s[4:6]}-{s[6:8]}" if len(s) == 8 else s


def build_row(item, company, fin, today):
    """공시 1건 → 엑셀 한 줄."""
    report_nm = " ".join((item.get("report_nm") or "").split())
    status = guess_status(report_nm)
    rcept_no = item.get("rcept_no", "")
    corp_name = company.get("corp_name") or item.get("corp_name", "")
    induty, induty_cat, induty_score = classify_induty(company.get("induty_code"))
    cats, score, reason = decide_category(corp_name, induty, induty_cat, induty_score)
    date = fmt_date(item.get("rcept_dt"))

    viewer = f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept_no}"
    note = f"[관련도 {score}/3] {report_nm} | 원문: {viewer}"

    return {
        "수집일자": today,
        "회사명": corp_name,
        "사업자번호": company.get("bizr_no", ""),
        "대표자명": company.get("ceo_nm", ""),
        "관할법원": "공시원문 확인",
        "사건번호": "공시원문 확인",
        "신청일": date if status in ("신청", "확인필요") else "",
        "결정일": date if status not in ("신청", "확인필요") else "",
        "사건상태": status,
        "면허/업종": induty,
        "분류": cats,
        "판단근거": reason,
        "자본금": fmt_won(fin.get("자본금")) if fin else "",
        "채무액": fmt_won(fin.get("부채총계")) if fin else "",
        "소재지": company.get("adres", ""),
        "전화번호": company.get("phn_no", ""),
        "데이터출처": f"DART {rcept_no}",
        "비고": note,
        "_score": score,
        "_rcept": rcept_no,
    }


# ─────────────────────────────────────────────────────────
# 출력
# ─────────────────────────────────────────────────────────

def write_excel(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "회생공시"

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    hit_fill = PatternFill("solid", fgColor="FFF2CC")   # 전기 관련 강조

    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in rows:
        ws.append([row.get(c, "") for c in COLUMNS])
        if row.get("_score", 0) >= 2:
            for cell in ws[ws.max_row]:
                cell.fill = hit_fill

    widths = {"회사명": 24, "소재지": 40, "비고": 70, "면허/업종": 24,
              "사업자번호": 14, "전화번호": 15, "데이터출처": 20}
    for idx, name in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(name, 12)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{ws.max_row}"

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ─────────────────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────────────────

def run_test(key):
    log("DART API 연결 확인 중…")
    data, err = call_api("company.json", {"corp_code": "00103662"}, key)
    if err:
        log(f"\n❌ 실패: {err}")
        if "IP" in err:
            log("   → DART 사이트에서 이 서버 IP를 등록했는지 확인하세요.")
        elif "등록" in err or "사용할 수 없" in err:
            log("   → https://opendart.fss.or.kr 에서 키를 다시 확인하세요.")
        return 1
    log("\n✅ 정상 — 테스트 조회 결과:")
    log(f"   회사명   : {data.get('corp_name')}")
    log(f"   대표자   : {data.get('ceo_nm')}")
    log(f"   사업자번호: {data.get('bizr_no')}")
    log(f"   전화     : {data.get('phn_no')}")
    log("\n이제 인자 없이 실행하면 수집이 시작됩니다: python3 collect_rehab.py")
    return 0


def main():
    p = argparse.ArgumentParser(description="DART 회생공시 수집기")
    p.add_argument("--key", help="DART API 키 (미지정 시 DART_API_KEY 환경변수)")
    p.add_argument("--days", type=int, default=30, help="최근 N일 (기본 30)")
    p.add_argument("--from", dest="bgn", help="시작일 YYYYMMDD")
    p.add_argument("--to", dest="end", help="종료일 YYYYMMDD")
    p.add_argument("--no-financials", action="store_true", help="자본금·부채 조회 생략")
    p.add_argument("--all", action="store_true", help="이미 본 건도 다시 포함")
    p.add_argument("--test", action="store_true", help="API 키 확인만")
    args = p.parse_args()

    key = get_api_key(args.key)
    if not key:
        log("❌ DART API 키가 없습니다.\n")
        log("  다음 중 하나로 지정하세요:")
        log('    export DART_API_KEY="키"')
        log("    또는 dart_key.txt 파일에 키만 저장")
        log("    또는 --key 옵션\n")
        log("  키 발급(무료): https://opendart.fss.or.kr")
        return 1

    if args.test:
        return run_test(key)

    today = datetime.now()
    end_de = args.end or today.strftime("%Y%m%d")
    bgn_de = args.bgn or (today - timedelta(days=args.days)).strftime("%Y%m%d")

    log("=" * 62)
    log("  트랙 A — DART 회생공시 수집")
    log("=" * 62)
    log(f"  기간   : {fmt_date(bgn_de)} ~ {fmt_date(end_de)}")
    log(f"  키워드 : {', '.join(KEYWORDS)}")
    log("")

    log("[1/4] 공시 검색")
    items = fetch_disclosures(key, bgn_de, end_de)
    if not items:
        log("\n해당 기간에 회생 관련 공시가 없습니다.")
        return 0
    log(f"  → {len(items)}건 발견\n")

    seen = set() if args.all else set(load_json(SEEN_FILE, []))
    new_items = [i for i in items if i.get("rcept_no") not in seen]
    if not new_items:
        log(f"신규 건이 없습니다. (기존 {len(items)}건은 이미 수집됨)")
        log("전부 다시 보려면 --all 을 붙이세요.")
        return 0
    log(f"[2/4] 신규 {len(new_items)}건 (기존 {len(items) - len(new_items)}건 제외)\n")

    log("[3/4] 기업정보 조회")
    cache = load_json(CACHE_FILE, {})
    rows = []
    for n, item in enumerate(new_items, 1):
        corp_code = item.get("corp_code", "")
        company = fetch_company(key, corp_code, cache)
        if "_error" in company:
            company = {"corp_name": item.get("corp_name", "")}

        fin = {}
        if not args.no_financials and corp_code:
            fin = fetch_financials(key, corp_code, today.year - 1)

        row = build_row(item, company, fin, today.strftime("%Y-%m-%d"))
        rows.append(row)

        mark = "★" if row["_score"] >= 2 else " "
        log(f"  {mark} [{n}/{len(new_items)}] {row['회사명']} — "
            f"{row['사건상태']} / {row['면허/업종']}")

    save_json(CACHE_FILE, cache)

    # 전기 관련도 높은 순 → 최신순
    rows.sort(key=lambda r: (-r["_score"], r["데이터출처"]), reverse=False)

    log("\n[4/4] 엑셀 저장")
    out_path = OUT_DIR / f"회생신청_정보_{today.strftime('%Y%m%d')}.xlsx"
    write_excel(rows, out_path)

    save_json(SEEN_FILE, sorted(seen | {i["rcept_no"] for i in new_items}))

    hits = [r for r in rows if r["_score"] >= 2]
    log(f"  → {out_path}")
    log("")
    log("=" * 62)
    log(f"  신규 공시 {len(rows)}건 저장 | 해당 업종 공시 {len(hits)}건")
    log("=" * 62)
    # 같은 회사가 공시 건수만큼 반복되므로 회사명으로 묶는다 (2026-08-21)
    by_corp = {}
    for r in hits:
        by_corp.setdefault(r["회사명"], []).append(r)

    if by_corp:
        log("\n  ★ 건설·전기·소방·설비통신 업체 %d곳" % len(by_corp))
        for corp, group in sorted(by_corp.items(), key=lambda kv: -kv[1][0]["_score"]):
            r = group[0]
            states = "/".join(dict.fromkeys(g["사건상태"] for g in group))
            more = " (공시 %d건)" % len(group) if len(group) > 1 else ""
            log(f"     · {corp} [{r.get('분류') or '분류없음'}] {states}{more}")
            log(f"       {r['면허/업종']}")
            log(f"       {r['소재지']} / {r['전화번호']}")
    log("")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        log("\n중단되었습니다.")
        sys.exit(130)
