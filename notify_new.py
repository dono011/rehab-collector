# -*- coding: utf-8 -*-
"""새로 뜬 전기 관련 회생 매각공고만 텔레그램으로 알린다.

만든 이유 (2026-08-21):
  수집기를 매일 자동으로 돌려도 결과가 VPS 엑셀로만 쌓여서
  새 매물이 떠도 사장님이 알 방법이 없었다.

언제 알리나:
  1) 새 공고가 잡혔을 때          ← 본래 목적
  2) 수집기가 실패했을 때          ← 조용히 죽는 것을 막는다
조용하면 정상이다. 매일 "이상 없음"을 보내지 않는다.

발송은 hermes send CLI로 그대로 전달한다. AI가 문장을 새로 짓지 않는다.
"""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
TRACK_B = os.path.join(HERE, "track_b_court")
SEEN = os.path.join(TRACK_B, "output", "seen.json")
DETAIL_DIR = os.path.join(TRACK_B, "output", "detail")
HERMES = "/usr/local/lib/hermes-agent/venv/bin/hermes"
PY = "/usr/bin/python3"


def read_seen():
    try:
        with open(SEEN, encoding="utf-8") as f:
            return set(json.load(f))
    except Exception:
        return set()


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


def company_of(key):
    """seen.json 열쇠 'site|id|회사명' 에서 회사명만 꺼낸다."""
    parts = key.split("|")
    return parts[-1].strip() if parts else key


def detail_of(name):
    """공고 원문에서 업종·자본금·연락처를 꺼낸다. 없으면 빈 값."""
    import re
    path = os.path.join(DETAIL_DIR, name.replace(" ", "_").replace("/", "_") + ".txt")
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8", errors="replace") as f:
            t = re.sub(r"\s+", " ", f.read())
    except Exception:
        return {}
    out = {}
    for label, key in (("업종", "업종"), ("납입자본금", "자본금"), ("관할법원", "법원")):
        m = re.search(
            label + r"\s(.{0,80}?)(?=\s(?:회사|관할법원|업종|회생절차|상장여부|납입자본금|홈페이지|주주의)\s)",
            t,
        )
        if m:
            out[key] = m.group(1).strip()
    m = re.search(r"(0\d{1,2}[-)]\s?\d{3,4}-\d{4})", t)
    if m:
        out["연락처"] = m.group(1)
    return out


def main():
    before = read_seen()

    r = subprocess.run(
        [PY, "collect_court.py", "--pages", "12", "--elec", "--detail"],
        cwd=TRACK_B, capture_output=True, text=True, timeout=1800,
    )
    sys.stdout.write(r.stdout or "")
    sys.stdout.write(r.stderr or "")

    if r.returncode != 0:
        tail = ((r.stdout or "") + (r.stderr or ""))[-500:]
        send_telegram("[회생매물 수집] 실패했습니다.\n종료코드 %s\n%s" % (r.returncode, tail))
        return r.returncode

    after = read_seen()
    new = sorted(after - before)
    if not new:
        print("새 공고 없음 — 알리지 않는다")
        return 0

    lines = ["[회생매물] 새 전기 관련 매각공고 %d건" % len(new), ""]
    for key in new:
        name = company_of(key)
        d = detail_of(name)
        lines.append("· %s" % name)
        if d.get("법원"):
            lines.append("  법원: %s" % d["법원"])
        if d.get("업종"):
            lines.append("  업종: %s" % d["업종"])
        if d.get("자본금"):
            lines.append("  자본금: %s" % d["자본금"])
        if d.get("연락처"):
            lines.append("  연락처: %s" % d["연락처"])
        lines.append("")
    lines.append("전기공사업 면허 실제 보유 여부는 확인이 필요합니다.")
    text = "\n".join(lines)

    ok, msg = send_telegram(text)
    print("텔레그램 발송:", "성공" if ok else "실패 " + msg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
