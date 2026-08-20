# -*- coding: utf-8 -*-
"""
트랙 B — 공고 파싱 모듈

법원 공고 목록 페이지의 HTML을 표준 레코드로 변환한다.
사이트 구조가 확정되면 손대는 곳은 이 파일 하나다.

레코드 형식:
    {법원, 사건번호, 회사명, 제목, 날짜, 링크}
"""

import re
from urllib.parse import urljoin

from bs4 import BeautifulSoup

# ─────────────────────────────────────────────────────────
# 정규식
# ─────────────────────────────────────────────────────────

# 회생·파산 사건번호   예) 2026회합163, 2026간회합168, 2025하합12
CASE_RE = re.compile(r"(20\d{2})\s?(간회합|회합|회단|하합|하단|개회)\s?(\d+)")

# 날짜  2026-08-19 / 2026.08.19 / 2026년 8월 19일
DATE_RE = re.compile(
    r"(20\d{2})[.\-/년]\s?(\d{1,2})[.\-/월]\s?(\d{1,2})"
)

# 법인 접미/접두 표기
CORP_TOKENS = ["주식회사", "(주)", "㈜", "유한회사", "(유)", "합자회사", "합명회사"]

# 공고 종류 → 사건상태
STATUS_MAP = [
    ("개시결정", "개시"), ("개시 결정", "개시"),
    ("개시신청", "신청"), ("개시 신청", "신청"),
    ("인가결정", "인가"), ("인가 결정", "인가"), ("계획인가", "인가"),
    ("종결결정", "종결"), ("종결", "종결"),
    ("폐지결정", "폐지"), ("폐지", "폐지"),
    ("기각", "기각"), ("취하", "취하"),
    ("매각공고", "매각"), ("매각", "매각"), ("인수의향", "매각"),
    ("채권신고", "채권신고"), ("파산선고", "파산"),
]

# 전기 관련 회사명 키워드 → 가중치
#
# 2026-08-20 실제 수집 결과(311건)를 보고 조정했다.
#   3점: 전기공사업일 가능성이 높은 어휘
#   2점: 전기 계열이지만 공사업이 아닐 수 있는 어휘 (제조·발전 등)
#   1점: 약한 신호
ELEC_KEYWORDS = [
    # 3점 — 전기공사·송배전 계열
    ("전기공사", 3), ("전기설비", 3), ("전설", 3), ("전기", 3),
    ("전력", 3), ("배전", 3), ("송전", 3), ("변전", 3), ("계전", 3),
    # 2점 — 전기 계열 (제조·발전·조명 포함)
    ("파워", 2), ("power", 2), ("일렉트", 2), ("electric", 2),
    ("케이블", 2), ("에너지", 2), ("발전", 2),
    ("솔라", 2), ("solar", 2), ("조명", 2), ("라이텍", 2),
    # 1점 — 약한 신호
    ("전자", 1), ("전공", 1), ("elec", 1),
]

# 회생 관련 공고만 남기기 위한 키워드
REHAB_KEYWORDS = ["회생", "파산", "매각", "인수의향", "M&A"]


# ─────────────────────────────────────────────────────────
# 필드 추출
# ─────────────────────────────────────────────────────────

def extract_case_no(text):
    """제목에서 사건번호를 뽑는다. 없으면 None."""
    m = CASE_RE.search(text or "")
    return f"{m.group(1)}{m.group(2)}{m.group(3)}" if m else None


def extract_date(text):
    """제목/셀에서 날짜를 YYYY-MM-DD 로. 없으면 빈 문자열."""
    m = DATE_RE.search(text or "")
    if not m:
        return ""
    y, mo, d = m.groups()
    return f"{y}-{int(mo):02d}-{int(d):02d}"


# 공고 문구 — 회사명에 섞이면 안 되는 단어
NOISE_WORDS = (
    "회생절차", "회생계획", "회생회사", "회생사건", "개시결정", "개시신청",
    "인가결정", "종결결정", "폐지결정", "파산선고", "채권신고기간", "채권신고",
    "매각공고", "재고자산", "유상증자", "제3자", "배정", "인수의향서", "인수의향",
    "공고", "안내", "결정", "신청", "기간", "사건", "에 대한", "에대한", "관련",
)

# 회사명 후보에서 법인 표기를 찾는 패턴 (접두형 / 접미형 모두)
NAME_PATTERNS = [
    (re.compile(r"주식회사\s*([^\s|│/,()]{2,20})"), "주식회사 {}"),
    (re.compile(r"([^\s|│/,()]{2,20})\s*주식회사"), "{} 주식회사"),
    (re.compile(r"유한회사\s*([^\s|│/,()]{2,20})"), "유한회사 {}"),
    (re.compile(r"([^\s|│/,()]{2,20})\s*유한회사"), "{} 유한회사"),
    (re.compile(r"㈜\s*([^\s|│/,()]{2,20})"), "㈜{}"),
    (re.compile(r"([^\s|│/,()]{2,20})\s*㈜"), "{}㈜"),
    (re.compile(r"\(주\)\s*([^\s|│/,()]{2,20})"), "(주){}"),
    (re.compile(r"([^\s|│/,()]{2,20})\s*\(주\)"), "{}(주)"),
]


CITIES = ("서울", "수원", "부산", "대구", "대전", "광주", "인천", "의정부",
          "창원", "청주", "전주", "춘천", "제주", "울산")

# 홀로 떨어져 있는 지명 — 공고 제목의 관할 표기일 뿐 회사명이 아니다.
# "부산전기(주)"처럼 붙어 있는 경우는 건드리지 않는다.
LONE_CITY_RE = re.compile(
    r"(?:(?<=^)|(?<=[\s|│/(),]))(" + "|".join(CITIES) + r")(?=[\s|│/(),]|$)")


def _strip_noise(text):
    """사건번호·날짜·법원명·관할지명·공고문구를 걷어낸다."""
    out = CASE_RE.sub(" ", text or "")
    out = DATE_RE.sub(" ", out)
    out = re.sub(r"(" + "|".join(CITIES) + r")\s?(회생법원|지방법원|지법)", " ", out)
    for w in NOISE_WORDS:
        out = out.replace(w, " ")
    out = LONE_CITY_RE.sub(" ", out)
    return " ".join(out.split())


def _is_plausible_name(name):
    """추출 결과가 회사명으로 볼 만한지."""
    if not name:
        return False
    core = re.sub(r"주식회사|유한회사|\(주\)|㈜|\s", "", name)
    if len(core) < 2 or len(core) > 30:
        return False
    if not re.search(r"[가-힣A-Za-z]", core):       # 글자가 없으면 탈락
        return False
    if re.fullmatch(r"[\d\W_]+", core):            # 숫자·기호뿐이면 탈락
        return False
    if any(w in name for w in ("안내", "공고", "기간", "결정", "신청")):
        return False
    if core in CITIES:                              # 관할 지명만 남은 경우
        return False
    return True


def extract_company(text):
    """
    공고 제목에서 회사명을 추출한다.

    실제 제목 형태가 다양하므로 여러 전략을 순서대로 시도하고,
    마지막에 타당성 검사를 통과한 것만 돌려준다.

      "회생절차개시결정 공고|수원|2026회합168|주식회사에이치엔티"
      "주식회사 ○○전기 회생절차개시결정"
      "○○전기(주)에 대한 회생절차 개시결정"
      "재고자산 매각공고 - 회생회사 한빛전기 주식회사"
    """
    if not text:
        return ""
    text = " ".join(text.split())

    candidates = []

    # 전략 1) 구분자로 나뉜 마지막 조각 (게시판 제목에서 가장 흔한 형태)
    for sep in ("|", "│"):
        if sep in text:
            tail = text.split(sep)[-1].strip(" -–—")
            if tail and not CASE_RE.search(tail):
                candidates.append(_strip_noise(tail) or tail)
            break

    # 전략 2) 법인 표기 패턴 (접두형·접미형 모두)
    cleaned = _strip_noise(text)
    for pattern, template in NAME_PATTERNS:
        m = pattern.search(cleaned)
        if m:
            candidates.append(template.format(m.group(1).strip(" -–—")))

    # 전략 3) 잡음을 모두 걷어내고 남은 것
    if cleaned:
        candidates.append(cleaned.strip(" -–—")[:40])

    for name in candidates:
        name = (name or "").strip(" -–—·")
        if _is_plausible_name(name):
            return name
    return ""


def extract_court(text):
    """제목/셀에서 법원명을 찾는다."""
    m = re.search(r"(서울|수원|부산|대구|대전|광주|인천|의정부|창원|청주|전주|"
                  r"춘천|제주|울산)\s?(회생법원|지방법원|지법)?", text or "")
    if not m:
        return ""
    city = m.group(1)
    kind = m.group(2) or ""
    return f"{city}{kind}" if kind else f"{city}회생법원"


def guess_status(title):
    """공고 제목 → 사건상태."""
    t = title or ""
    for needle, label in STATUS_MAP:
        if needle in t:
            return label
    return "확인필요"


def elec_score(company, title=""):
    """
    회사명 기반 전기 관련도 (0~3).

    트랙 A는 DART 업종코드를 쓰지만, 비상장사는 업종 정보가 없어
    회사명 키워드로 1차 선별한다. 정확도가 낮으므로
    최종 확인은 반드시 한국전기공사협회 조회로 해야 한다.
    """
    haystack = f"{company} {title}".lower()
    best = 0
    for kw, weight in ELEC_KEYWORDS:
        if kw.lower() in haystack:
            best = max(best, weight)
    return best


def is_rehab_notice(title):
    """
    회생·파산·매각 관련 공고인지.

    법원 이름 자체에 '회생'이 들어가므로("서울회생법원"), 이를 먼저
    걷어낸 뒤 판정한다. 그러지 않으면 회생법원이 올린 모든 공고가
    통과해 버린다.
    """
    text = (title or "")
    text = re.sub(r"(서울|수원|부산|대구|대전|광주|인천|의정부|창원|청주|전주|"
                  r"춘천|제주|울산)?회생법원", " ", text)
    text = text.lower()
    return any(kw.lower() in text for kw in REHAB_KEYWORDS)


# ─────────────────────────────────────────────────────────
# 페이지 파서
# ─────────────────────────────────────────────────────────

def parse_table(html, base_url=""):
    """
    <table> 기반 목록 페이지를 파싱한다.
    한국 공공기관 게시판 대부분이 이 구조다.
    """
    soup = BeautifulSoup(html, "html.parser")
    records = []

    for table in soup.find_all("table"):
        for tr in table.find_all("tr"):
            cells = tr.find_all(["td", "th"])
            if len(cells) < 2:
                continue

            texts = [" ".join(c.get_text(" ", strip=True).split()) for c in cells]
            joined = " ".join(texts)
            if not joined.strip():
                continue

            link = ""
            a = tr.find("a", href=True)
            if a:
                href = a["href"].strip()
                if href and not href.lower().startswith("javascript"):
                    link = urljoin(base_url, href) if base_url else href

            # 제목: 가장 긴 셀 (대부분 제목 칸)
            title = max(texts, key=len) if texts else ""
            if a:
                a_text = " ".join(a.get_text(" ", strip=True).split())
                if len(a_text) > 4:
                    title = a_text

            rec = _build_record(title, joined, link)
            if rec:
                records.append(rec)

    return records


def parse_list(html, base_url=""):
    """<ul>/<ol>/<div> 카드 형태 목록 페이지를 파싱한다."""
    soup = BeautifulSoup(html, "html.parser")
    records = []

    for a in soup.find_all("a", href=True):
        text = " ".join(a.get_text(" ", strip=True).split())
        if len(text) < 6:
            continue

        href = a["href"].strip()
        if href.lower().startswith("javascript"):
            href = ""
        link = urljoin(base_url, href) if (href and base_url) else href

        # 주변 텍스트까지 포함해 날짜/법원을 찾는다
        parent = a.find_parent(["li", "tr", "div"])
        context = " ".join(parent.get_text(" ", strip=True).split()) if parent else text

        rec = _build_record(text, context, link)
        if rec:
            records.append(rec)

    return records


def _build_record(title, context, link):
    """
    제목+주변텍스트 → 레코드. 회생 공고가 아니면 None.

    판정은 제목을 기준으로 한다. 주변 텍스트에는 게시판 머리글이나
    다른 행의 내용이 섞여 들어오므로, 사건번호가 함께 있을 때만
    보조 근거로 인정한다.
    """
    case_no = extract_case_no(title) or extract_case_no(context)

    if not is_rehab_notice(title):
        if not (case_no and is_rehab_notice(context)):
            return None

    company = extract_company(title)
    if not company and not case_no:
        return None

    return {
        "법원": extract_court(title) or extract_court(context),
        "사건번호": case_no or "",
        "회사명": company,
        "제목": title,
        "날짜": extract_date(title) or extract_date(context),
        "링크": link,
        "상태": guess_status(title) or guess_status(context),
        "관련도": elec_score(company, title),
    }


def dedupe(records):
    """사건번호+제목 기준 중복 제거. 순서는 유지."""
    seen = set()
    result = []
    for r in records:
        key = (r.get("사건번호", ""), r.get("제목", "")[:60])
        if key in seen:
            continue
        seen.add(key)
        result.append(r)
    return result


def parse_auto(html, base_url=""):
    """표 파서를 먼저 쓰고, 결과가 없으면 목록 파서로 넘어간다."""
    records = parse_table(html, base_url)
    if not records:
        records = parse_list(html, base_url)
    return dedupe(records)


# ─────────────────────────────────────────────────────────
# 대법원 M&A 공고게시판 전용 파서
#
# 2026-08-20 실제 페이지 구조 확인 후 작성.
#   https://www.scourt.go.kr/portal/notice/mainfo/MaNoticeList.work
#
# 표 구조가 고정되어 있고 회사명이 별도 칼럼으로 제공되므로,
# 범용 파서의 회사명 추측 로직이 필요 없다.
#
#   번호 | 관할법원 | 업종 | 회사 | 작성일
#   345  | 수원회생법원 | 수원회생법원-기타-㈜이노피아테크 기타 | ㈜이노피아테크 | 2026-08-18
# ─────────────────────────────────────────────────────────

# business 파라미터 코드 → 업종명 (실제 페이지에서 확인한 전체 17개)
BUSINESS_CODES = {
    "01": "제조업(기계,금속,중공업)",
    "02": "제조업(전기전자)",
    "03": "제조업(석유,화학)",
    "04": "제조업(섬유)",
    "05": "제조업(가구,건설ㆍ건축자재)",
    "06": "제조업(문구,완구,서적,화장지류)",
    "08": "건설,엔지니어링,설계",
    "09": "가스,에너지,수도",
    "10": "약품",
    "11": "식품,농축수산",
    "12": "무역",
    "13": "유통,백화점",
    "14": "의류,패션",
    "15": "운수,창고",
    "16": "관광,숙박,레저,서비스",
    "17": "기타",
}

# 전기공사업이 속할 가능성이 있는 업종 → 가중치
#
# ⚠️ "건설,엔지니어링,설계"(08)를 높게 잡으면 삼부토건·경남기업 같은
#    종합건설사가 무더기로 딸려 온다. 실제 수집 결과 51건 중 20건 가까이가
#    그런 잡음이었다. 그래서 08은 참고 수준(1점)으로 낮춘다.
#    전기공사업체는 대개 회사명에 신호가 있으므로 이름 쪽에서 잡힌다.
BUSINESS_SCORE = {
    "02": 2,   # 제조업(전기전자)
    "09": 2,   # 가스,에너지,수도
    "08": 1,   # 건설,엔지니어링,설계 — 종합건설이 대다수
}

# bub_cd 파라미터 코드 → 법원명 (실제 페이지에서 확인한 전체)
COURT_CODES = {
    "000210": "서울중앙지방법원", "000221": "서울회생법원",
    "000214": "의정부지방법원",   "000240": "인천지방법원",
    "000250": "수원지방법원",     "000249": "수원회생법원",
    "000260": "춘천지방법원",     "000280": "대전지방법원",
    "000291": "대전회생법원",     "000270": "청주지방법원",
    "000310": "대구지방법원",     "000321": "대구회생법원",
    "000410": "부산지방법원",     "000443": "부산회생법원",
    "000411": "울산지방법원",     "000420": "창원지방법원",
    "000510": "광주지방법원",     "000543": "광주회생법원",
    "000520": "전주지방법원",     "000530": "제주지방법원",
}


def _split_induty(cell, court, company):
    """
    업종 칸에서 업종명만 뽑는다.

    칸 내용이 "{법원}-{업종}-{회사} {업종}" 형태로 뭉쳐 나오므로
    하이픈으로 끊어 가운데 조각을 쓴다. 업종명 자체에 쉼표는 있어도
    하이픈은 없어서 이 방법이 안전하다.
    """
    cell = " ".join((cell or "").split())
    if not cell:
        return ""
    parts = cell.split("-")
    if len(parts) >= 3:
        return parts[1].strip()
    # 형태가 다르면 법원명·회사명을 걷어내고 남은 것
    out = cell.replace(court or "", " ").replace(company or "", " ")
    return " ".join(out.split())


def parse_ma_notice(html, base_url=""):
    """
    M&A 공고게시판 목록 → 레코드 목록.

    범용 파서와 달리 칼럼 위치가 확정되어 있어 추측이 없다.
    """
    soup = BeautifulSoup(html, "html.parser")
    records = []

    for tr in soup.find_all("tr"):
        cells = [" ".join(td.get_text(" ", strip=True).split())
                 for td in tr.find_all("td")]
        if len(cells) < 5:
            continue

        no, court, induty_cell, company, date = cells[0], cells[1], cells[2], cells[3], cells[4]

        if not no.isdigit() or not company:
            continue                      # 머리글·페이지네이션 행

        induty = _split_induty(induty_cell, court, company)

        link = ""
        a = tr.find("a", href=True)
        if a:
            href = a["href"].strip()
            if href and not href.lower().startswith("javascript"):
                link = urljoin(base_url, href) if base_url else href

        # 회사명 신호와 업종 신호를 따로 매긴다.
        # 합쳐버리면 어느 쪽에서 걸린 건지 알 수 없어 잡음 판별이 안 된다.
        name_score = elec_score(company)
        induty_score = 0
        for code, weight in BUSINESS_SCORE.items():
            if BUSINESS_CODES[code] in induty:
                induty_score = max(induty_score, weight)

        records.append({
            "번호": no,
            "법원": court,
            "업종": induty,
            "회사명": company,
            "제목": f"{company} M&A 매각공고 ({induty})",
            "날짜": extract_date(date) or date,
            "링크": link,
            "사건번호": extract_case_no(induty_cell) or "",
            "상태": "매각",
            "관련도": max(name_score, induty_score),
            "이름점수": name_score,
            "업종점수": induty_score,
        })

    return records
