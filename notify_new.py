# -*- coding: utf-8 -*-
"""새로 뜬 건설·전기·소방 매물만 텔레그램으로 알린다.

만든 이유 (2026-08-21):
  수집기를 매일 자동으로 돌려도 결과가 서버 엑셀로만 쌓여
  새 매물이 떠도 사장님이 알 방법이 없었다.

무엇을 도나:
  트랙 B — 대법원 M&A 매각공고   (팔려고 내놓은 회사. 매각주간사 전화가 붙는다)
  트랙 C — 법원 회생·파산 공고    (절차가 열린 회사 전부. 파산은 관재인 전화가 붙는다)
  두 자료원은 서로 구멍을 메운다. 한쪽만 돌리면 놓친다.

언제 알리나:
  1) 새 공고가 잡혔을 때          <- 본래 목적
  2) 수집기가 실패했을 때          <- 조용히 죽는 것을 막는다
조용하면 정상이다. 매일 "이상 없음"을 보내지 않는다.

발송은 팩스 감시와 같은 hermes send CLI 를 쓴다. 헤르메스는 그대로 전달만 하고
문장을 새로 짓지 않는다 (지어내는 문제가 끼어들 여지를 없앤다).

사용법:
    python3 notify_new.py            # 수집 + 새 것 있으면 알림
    python3 notify_new.py --quiet    # 수집만 하고 알리지 않음 (처음 채울 때)
"""
import argparse
import json
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
TRACK_B = os.path.join(HERE, "track_b_court")
TRACK_A = os.path.join(HERE, "track_a_dart")
TRACK_C = os.path.join(HERE, "track_c_insolvency")
ENV_FILE = os.path.join(HERE, ".env")
HERMES = "/usr/local/lib/hermes-agent/venv/bin/hermes"
PY = "/usr/bin/python3"

sys.path.insert(0, HERE)
import categories as CAT  # noqa: E402

# 알릴 업종. 여기 없는 것은 잡혀도 알리지 않는다.
WANTED = ("건설", "전기", "소방", "설비통신")
MAX_LINES = 60  # 텔레그램 한 통이 너무 길어지지 않게


def send_telegram(text):
    """hermes send CLI로 텔레그램 홈채널에 전송. AI 판단 없이 그대로 전달."""
    try:
        r = subprocess.run(
            [HERMES, "send", "--to", "telegram", text],
            capture_output=True, text=True, timeout=60,
        )
        return r.returncode == 0, (r.stdout or "") + (r.stderr or "")
    except Exception as e:
        return False, str(e)


def run(cwd, argv, timeout=2400):
    r = subprocess.run([PY] + argv, cwd=cwd, capture_output=True,
                       text=True, timeout=timeout)
    sys.stdout.write(r.stdout or "")
    sys.stdout.write(r.stderr or "")
    return r


def read_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


# ── 트랙 B ────────────────────────────────────────────────
def track_b():
    """매각공고. seen.json 을 실행 전후로 비교해 새 것만 골라낸다."""
    seen_path = os.path.join(TRACK_B, "output", "seen.json")
    before = set(read_json(seen_path, []))

    r = run(TRACK_B, ["collect_court.py", "--pages", "12", "--detail"])
    if r.returncode != 0:
        return None, ((r.stdout or "") + (r.stderr or ""))[-400:]

    after = set(read_json(seen_path, []))
    items = []
    for key in sorted(after - before):
        name = key.split("|")[-1].strip()
        d = detail_of_b(name)
        c = CAT.classify(name, d.get("업종", ""))
        if not any(cat in WANTED for cat in c["카테고리"]):
            continue
        items.append({
            "회사": name, "분류": "/".join(c["카테고리"]),
            "법원": d.get("법원", ""), "업종": d.get("업종", ""),
            "자본금": d.get("자본금", ""), "연락처": d.get("연락처", ""),
            "구분": "매각공고",
        })
    return items, None


def detail_of_b(name):
    """매각공고 원문에서 업종·자본금·연락처를 꺼낸다. 없으면 빈 값."""
    path = os.path.join(TRACK_B, "output", "detail",
                        name.replace(" ", "_").replace("/", "_") + ".txt")
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            t = re.sub(r"\s+", " ", f.read())
    except Exception:
        return {}
    out = {}
    stop = r"(?=\s(?:회사|관할법원|업종|회생절차|상장여부|납입자본금|홈페이지|주주의)\s)"
    for label, key in (("업종", "업종"), ("납입자본금", "자본금"), ("관할법원", "법원")):
        m = re.search(label + r"\s(.{0,80}?)" + stop, t)
        if m:
            out[key] = m.group(1).strip()
    m = re.search(r"(0\d{1,2}[-)]\s?\d{3,4}-\d{4})", t)
    if m:
        out["연락처"] = m.group(1)
    return out


# ── 트랙 A ────────────────────────────────────────────────
def load_env():
    """.env 의 DART_API_KEY 를 읽는다. 없으면 None."""
    try:
        with open(ENV_FILE, encoding="utf-8") as f:
            for line in f:
                if line.startswith("DART_API_KEY="):
                    v = line.split("=", 1)[1].strip()
                    return v or None
    except Exception:
        pass
    return None


def track_a():
    """DART 공시. 법원 자료에 없는 사업자번호·대표자·전화·자본금을 채워 준다.
    상장사·공시대상 법인만 나오므로 건수는 적다."""
    key = load_env()
    if not key:
        return [], None          # 키가 없으면 조용히 건너뛴다 (실패 아님)

    seen_path = os.path.join(TRACK_A, "output", "seen.json")
    before = set(read_json(seen_path, []))

    env = dict(os.environ, DART_API_KEY=key)
    r = subprocess.run([PY, "collect_rehab.py", "--days", "30"], cwd=TRACK_A,
                       capture_output=True, text=True, timeout=2400, env=env)
    sys.stdout.write(r.stdout or "")
    sys.stdout.write(r.stderr or "")
    if r.returncode != 0:
        return None, ((r.stdout or "") + (r.stderr or ""))[-400:]

    after = set(read_json(seen_path, []))
    if not (after - before):
        return [], None

    # 새로 들어온 접수번호에 해당하는 줄만 엑셀에서 꺼낸다
    import glob
    files = sorted(glob.glob(os.path.join(TRACK_A, "output", "회생신청_정보_*.xlsx")))
    if not files:
        return [], None
    try:
        from openpyxl import load_workbook
        ws = load_workbook(files[-1]).active
    except Exception as e:
        return None, "엑셀 읽기 실패: %s" % e

    head = [c.value for c in ws[1]]
    idx = dict((n, i) for i, n in enumerate(head))
    fresh = after - before
    by_corp = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        src = str(row[idx.get("데이터출처", 0)] or "")
        rcept = src.replace("DART ", "").strip()
        if rcept not in fresh:
            continue
        cats = [c for c in str(row[idx.get("분류", 0)] or "").split("/") if c]
        if not any(c in WANTED for c in cats):
            continue
        corp = str(row[idx.get("회사명", 0)] or "")
        if corp in by_corp:                      # 같은 회사 여러 공시 → 한 줄
            continue
        by_corp[corp] = {
            "회사": corp, "분류": "/".join(cats),
            "법원": "", "업종": str(row[idx.get("면허/업종", 0)] or ""),
            "상태": str(row[idx.get("사건상태", 0)] or ""),
            "자본금": str(row[idx.get("자본금", 0)] or ""),
            "주소": str(row[idx.get("소재지", 0)] or ""),
            "연락처": str(row[idx.get("전화번호", 0)] or ""),
            "사업자번호": str(row[idx.get("사업자번호", 0)] or ""),
            "구분": "DART공시",
        }
    return list(by_corp.values()), None


# ── 트랙 C ────────────────────────────────────────────────
def track_c():
    """회생·파산 공고. 수집기가 신규.json 을 직접 써 준다."""
    new_path = os.path.join(TRACK_C, "output", "신규.json")
    if os.path.exists(new_path):
        os.remove(new_path)

    r = run(TRACK_C, ["collect_insolvency.py"])
    if r.returncode != 0:
        return None, ((r.stdout or "") + (r.stderr or ""))[-400:]

    items = []
    for rec in read_json(new_path, []):
        cats = [c for c in (rec.get("분류") or "").split("/") if c]
        if not any(c in WANTED for c in cats):
            continue
        items.append({
            "회사": rec.get("채무자명", ""), "분류": rec.get("분류", ""),
            "법원": rec.get("법원명", ""), "업종": rec.get("사건구분", ""),
            "사건번호": rec.get("사건번호", ""), "공고": rec.get("공고제목", ""),
            "주소": rec.get("회사주소", ""),
            "관재인": rec.get("관리인_관재인", ""),
            "연락처": rec.get("관재인_전화", ""),
            "구분": "회생·파산공고",
        })
    return items, None


# ── 알림 문구 ──────────────────────────────────────────────
def build_message(items):
    by_cat = {}
    for it in items:
        head = (it["분류"].split("/")[0] if it["분류"] else "기타")
        by_cat.setdefault(head, []).append(it)

    lines = ["[매물알림] 새 공고 %d건" % len(items), ""]
    for cat in WANTED:
        group = by_cat.get(cat)
        if not group:
            continue
        lines.append("■ %s (%d건)" % (cat, len(group)))
        for it in group:
            lines.append("· %s [%s]" % (it["회사"], it["구분"]))
            bits = [b for b in (it.get("법원"), it.get("업종"),
                                it.get("사건번호"), it.get("공고")) if b]
            if bits:
                lines.append("  " + " / ".join(bits))
            if it.get("자본금"):
                lines.append("  자본금: %s" % it["자본금"])
            if it.get("주소"):
                lines.append("  주소: %s" % it["주소"])
            if it.get("사업자번호"):
                lines.append("  사업자번호: %s" % it["사업자번호"])
            if it.get("상태"):
                lines.append("  상태: %s" % it["상태"])
            if it.get("관재인"):
                lines.append("  관재인: %s" % it["관재인"])
            if it.get("연락처"):
                lines.append("  연락처: %s" % it["연락처"])
            else:
                lines.append("  연락처: 공고에 없음")
        lines.append("")

    if len(lines) > MAX_LINES:
        lines = lines[:MAX_LINES] + ["", "(너무 길어 줄임 — 서버 엑셀에 전부 있습니다)"]
    lines.append("면허 실제 보유 여부는 확인이 필요합니다. 이름으로 추린 결과입니다.")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--quiet", action="store_true",
                    help="수집만 하고 알리지 않음 (처음 채울 때)")
    args = ap.parse_args()

    items, failures = [], []

    for label, fn in (("매각공고", track_b), ("회생·파산공고", track_c),
                      ("DART공시", track_a)):
        print("\n===== %s =====" % label)
        try:
            got, err = fn()
        except Exception as e:
            got, err = None, str(e)
        if err:
            failures.append("%s: %s" % (label, err))
        elif got:
            items.extend(got)

    if failures and not args.quiet:
        send_telegram("[매물알림] 수집 실패\n\n" + "\n\n".join(failures))

    if not items:
        print("\n새 공고 없음 — 알리지 않는다")
        return 1 if failures else 0

    print("\n새 공고 %d건" % len(items))
    if args.quiet:
        print("--quiet 이므로 알리지 않는다")
        return 0

    ok, msg = send_telegram(build_message(items))
    print("텔레그램 발송:", "성공" if ok else "실패 " + msg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
