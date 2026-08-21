# -*- coding: utf-8 -*-
"""트랙 C — 법원 회생·파산 공고 수집기 (건설·전기·소방·설비통신)

대법원 회생·파산 공고(ssgo.scourt.go.kr)를 훑어 건설 계열 법인만 골라내고,
각 건의 공고 원문을 열어 대표자·회사주소·파산관재인 연락처까지 붙인다.

컴1에서 사람이 직접 돌리던 court_electric_collect.py + court_electric_contacts.py 를
서버용 하나로 합친 것이다 (2026-08-21).

주의 1: 이 게시판은 법원이 업종을 매기지 않는다. 회사 이름밖에 없어서
        categories.py 의 이름 추측에 의존한다. 반드시 놓치고 반드시 잘못 잡는다.
주의 2: 회생 사건은 공고 원문에 관리인 연락처가 아예 없다(파산에만 있다). 원자료의 한계다.

사용법:
    python3 collect_insolvency.py              # 새로 뜬 것만 원문 열람
    python3 collect_insolvency.py --all        # 이미 본 것도 다시
    python3 collect_insolvency.py --no-detail  # 원문 열람 생략 (빠름)
"""
import argparse
import json
import os
import sys
import time
from datetime import datetime

import requests
from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import categories as CAT  # noqa: E402

OUT_DIR = os.path.join(HERE, "output")
SEEN_FILE = os.path.join(OUT_DIR, "seen.json")

LIST_URL = "https://ssgo.scourt.go.kr/ssgo/ssgo930/selectRhblBnkpPbancLst.on"
POPUP_URL = "https://ssgo.scourt.go.kr/ssgo/ssgo930/selectBfCsPbancPviewInf.on"
HEADERS = {
    "Content-Type": "application/json;charset=UTF-8",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer": "https://ssgo.scourt.go.kr/ssgo/ssgo930/rhblBnkp.on",
    "X-Requested-With": "XMLHttpRequest",
}

# 2026-08-21 --codes 로 확인한 전체 법원.
# 기존 컴1 스크립트는 12곳만 봤다 (제주·춘천·광주회생·서울중앙 등이 빠져 있었다).
COURTS = [
    ("000221", "서울회생법원"), ("000210", "서울중앙지방법원"),
    ("000249", "수원회생법원"), ("000250", "수원지방법원"),
    ("000214", "의정부지방법원"), ("000240", "인천지방법원"),
    ("000291", "대전회생법원"), ("000280", "대전지방법원"),
    ("000270", "청주지방법원"), ("000260", "춘천지방법원"),
    ("000321", "대구회생법원"), ("000310", "대구지방법원"),
    ("000443", "부산회생법원"), ("000410", "부산지방법원"),
    ("000411", "울산지방법원"), ("000420", "창원지방법원"),
    ("000543", "광주회생법원"), ("000510", "광주지방법원"),
    ("000520", "전주지방법원"), ("000530", "제주지방법원"),
]
# 일반회생(taskDvs=3)은 뺐다 — 개인이 신청하는 것이라 채무자명이 사람 이름이다.
# 2026-08-21 확인: 전국 일반회생 700여 건 중 회사명은 0건이었다. 법인 매물과 무관하다.
TASK_DVS = {"4": "법인회생", "5": "법인파산"}
PAGE_SIZE = 50
DELAY = 0.3


def log(msg):
    print(msg, flush=True)


def new_session():
    s = requests.Session()
    s.headers.update({"User-Agent": HEADERS["User-Agent"]})
    try:
        s.get("https://ssgo.scourt.go.kr/ssgo/ssgo930/rhblBnkp.on", timeout=15)
    except Exception as e:
        log("  세션 준비 경고: %s" % e)
    return s


def search_payload(task_dvs, cort_cd, page_no):
    return {
        "dma_search": {"taskDvs": task_dvs, "srchType": "pstgBgng", "cortCd": cort_cd,
                       "pstgDvs": "999", "csYr": "26", "csDvsCd": "253",
                       "csSrno": "", "csNo": "", "debtrNm": "", "jdbnCd": ""},
        "dma_pageInfo": {"pageNo": page_no, "pageSize": PAGE_SIZE, "bfPageNo": "",
                         "startRowNo": "", "totalCnt": "", "totalYn": "Y"},
    }


def fetch_page(s, task_dvs, cort_cd, page_no):
    r = s.post(LIST_URL, json=search_payload(task_dvs, cort_cd, page_no),
               headers=HEADERS, timeout=25)
    d = r.json().get("data", {})
    total = d.get("dma_pageInfo", {}).get("totalCnt", 0)
    return (d.get("dlt_pbancLst") or []), int(total or 0)


def get_detail(s, item, task_dvs):
    """공고 원문 팝업 → 대표자·주소·관재인. 회생 사건은 관재인 항목이 없다."""
    pp = {"dma_popupSearch": {"cortCd": item["cortCd"], "csNo": item["csNo"],
                              "inetPbancDvsCd": item["inetPbancDvsCd"],
                              "inetPbancSeq": item["inetPbancSeq"], "taskDvs": task_dvs}}
    m = s.post(POPUP_URL, json=pp, headers=HEADERS,
               timeout=25).json().get("data", {}).get("rtnMap", {})
    bt = m.get("btprtInf") or m.get("btprtDebtrInf") or {}
    if not bt:
        lst = m.get("btprtPtnrLst") or []
        bt = lst[0] if lst else {}
    mng_list = m.get("mngChargLst") or []
    mng = mng_list[0] if mng_list else {}
    return {
        "대표자": bt.get("rprsvNm", "") or "",
        "회사주소": bt.get("btprtAddr", "") or "",
        "관리인_관재인": mng.get("hlprNm", "") or "",
        "관재인_전화": mng.get("prcdRltnrTelno", "") or "",
        "관재인_주소": ((mng.get("basAddr", "") or "") + " "
                    + (mng.get("dtlAddr", "") or "")).strip(),
    }


def read_seen():
    try:
        with open(SEEN_FILE, encoding="utf-8") as f:
            return set(json.load(f))
    except Exception:
        return set()


def write_seen(keys):
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(SEEN_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(keys), f, ensure_ascii=False, indent=2)


def key_of(rec):
    return rec["채무자명"] + "|" + rec["사건번호"]


def collect(min_score=2):
    s = new_session()
    found, errors = [], []

    for cort_cd, cort_nm in COURTS:
        for task_dvs, task_nm in TASK_DVS.items():
            try:
                items, total = fetch_page(s, task_dvs, cort_cd, 1)
                if total == 0:
                    continue
                pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
                batch = list(items)
                for p in range(2, pages + 1):
                    time.sleep(DELAY)
                    more, _ = fetch_page(s, task_dvs, cort_cd, p)
                    batch.extend(more)

                hit = 0
                for it in batch:
                    name = it.get("btprtNm", "")
                    c = CAT.classify(name)
                    if c["점수"] < min_score:
                        continue
                    hit += 1
                    found.append({
                        "법원명": cort_nm, "사건구분": task_nm, "_taskDvs": task_dvs,
                        "채무자명": name,
                        "사건번호": it.get("csNoNm", ""),
                        "공고제목": it.get("pbancTitlNm", ""),
                        "공고시작일": it.get("pbancBgngYmd", ""),
                        "재판부": it.get("jdbnCdNm", ""),
                        "분류": "/".join(c["카테고리"]),
                        "관련도": c["점수"], "판단근거": c["근거"],
                        "_item": it,
                    })
                log("  %-14s %-6s %5d건 -> 해당 %d건" % (cort_nm, task_nm, total, hit))
            except Exception as e:
                msg = "%s %s: %s" % (cort_nm, task_nm, e)
                log("  오류 " + msg)
                errors.append(msg)
    return found, errors


def main():
    ap = argparse.ArgumentParser(description="법원 회생·파산 공고 수집기 (건설·전기·소방)")
    ap.add_argument("--all", action="store_true", help="이미 본 건도 포함")
    ap.add_argument("--no-detail", action="store_true", help="공고 원문 열람 생략")
    ap.add_argument("--min-score", type=int, default=2, help="관련도 최소 점수 (기본 2)")
    args = ap.parse_args()

    log("=" * 62)
    log(" 법원 회생·파산 공고 수집 — 건설·전기·소방·설비통신")
    log(" 법원 %d곳 / %s" % (len(COURTS), ", ".join(TASK_DVS.values())))
    log("=" * 62)

    found, errors = collect(args.min_score)
    log("")
    log("  해당 업종 %d건 (오류 %d건)" % (len(found), len(errors)))
    if errors and not found:
        log("  전부 실패했습니다.")
        return 2

    # 중복 제거 — 같은 회사+사건번호는 최신 공고 하나만 남긴다
    uniq = {}
    for r in sorted(found, key=lambda x: x["공고시작일"], reverse=True):
        uniq.setdefault(key_of(r), r)
    records = sorted(uniq.values(), key=lambda x: x["공고시작일"], reverse=True)
    log("  중복 제거 후 %d건" % len(records))

    seen = read_seen()
    new_only = [r for r in records if key_of(r) not in seen]
    log("  신규 %d건" % len(new_only))

    targets = records if args.all else new_only
    if not args.no_detail and targets:
        log("")
        log("  공고 원문 열람 (%d건)" % len(targets))
        s = new_session()
        for i, r in enumerate(targets, 1):
            try:
                r.update(get_detail(s, r["_item"], r["_taskDvs"]))
                log("   [%d/%d] %s — 관재인 %s %s" % (
                    i, len(targets), r["채무자명"],
                    r.get("관리인_관재인") or "없음", r.get("관재인_전화") or ""))
            except Exception as e:
                log("   [%d/%d] %s — 원문 실패: %s" % (i, len(targets), r["채무자명"], e))
            time.sleep(DELAY)

    os.makedirs(OUT_DIR, exist_ok=True)
    cols = ["법원명", "사건구분", "채무자명", "분류", "관련도", "판단근거", "사건번호",
            "공고제목", "공고시작일", "재판부", "대표자", "회사주소",
            "관리인_관재인", "관재인_전화", "관재인_주소"]
    wb = Workbook()
    ws = wb.active
    ws.title = "회생파산공고"
    ws.append(cols)
    for r in records:
        ws.append([str(r.get(c, "") or "") for c in cols])
    path = os.path.join(OUT_DIR,
                        "법원공고_회생파산_%s.xlsx" % datetime.now().strftime("%Y%m%d"))
    wb.save(path)
    log("")
    log("  저장: %s" % path)

    out = os.path.join(OUT_DIR, "신규.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump([{k: v for k, v in r.items() if not k.startswith("_")}
                   for r in new_only], f, ensure_ascii=False, indent=2)
    log("  신규 목록: %s (%d건)" % (out, len(new_only)))

    write_seen(seen | {key_of(r) for r in records})
    return 0


if __name__ == "__main__":
    sys.exit(main())
