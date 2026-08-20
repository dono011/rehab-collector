#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
트랙 B — 1단계: 법원 사이트 정찰 (Probe)

크롤러를 만들기 전에, 대상 사이트가 실제로 어떤 상태인지 확인한다.
수동 curl을 대신하는 도구로, 이 스크립트 하나만 VPS에서 돌리면
필요한 정보가 전부 리포트 파일로 정리된다.

확인하는 것:
  1. robots.txt      — 크롤링이 허용되는가
  2. HTTP 응답        — 접근이 되는가 (200/403/404)
  3. 렌더링 방식      — HTML에 데이터가 있는가, JS로 그리는가
  4. 사건번호 존재    — 2026회합163 같은 패턴이 실제로 보이는가
  5. 표 구조         — 파싱 가능한 <table>/<ul>이 있는가

사용법:
    python3 probe.py                  # 전체 정찰
    python3 probe.py --save-html      # 원본 HTML도 저장 (파서 개발용)
    python3 probe.py --url "주소"      # 특정 주소만 확인

결과:
    output/probe_report.txt   ← 이 파일 내용을 그대로 전달하면 됨
"""

import argparse
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

try:
    import requests
except ImportError:
    print("requests 가 필요합니다:  pip install requests --break-system-packages")
    sys.exit(1)

HERE = Path(__file__).resolve().parent
OUT_DIR = HERE / "output"
REPORT = OUT_DIR / "probe_report.txt"

UA = ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")

# 확인할 후보 주소. 일부는 추정이며, 정찰의 목적이 어느 것이 살아있는지 가리는 것.
TARGETS = [
    ("대법원 법원공고 목록",
     "https://www.scourt.go.kr/portal/notice/paper/paperList.work"),
    ("대법원 회생회사 M&A 안내  ★매각공고",
     "https://www.scourt.go.kr/portal/notice/mna/guide/index.html"),
    ("대법원 회생회사 M&A 공고목록 (추정)",
     "https://www.scourt.go.kr/portal/notice/mna/mnaList.work"),
    ("서울회생법원",
     "https://slb.scourt.go.kr/main/Main.work"),
    ("대한민국 관보",
     "https://gwanbo.go.kr/main.do"),
]

# 회생·파산 사건번호 패턴  예) 2026회합163, 2026간회합168, 2025하합12
CASE_RE = re.compile(r"20\d{2}\s?(?:간회합|회합|회단|하합|하단|개회|국승)\s?\d+")

# JS 프레임워크 흔적 → 서버가 빈 껍데기만 주는지 판단
JS_HINTS = ["ng-app", "v-app", "__NUXT__", "__NEXT_DATA__",
            "react-root", "vue.js", "angular.min.js"]


class Tee:
    """화면과 파일에 동시에 쓴다."""

    def __init__(self, path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.fh = path.open("w", encoding="utf-8")

    def __call__(self, msg=""):
        print(msg, flush=True)
        self.fh.write(msg + "\n")

    def close(self):
        self.fh.close()


def fetch(url, timeout=25):
    """(response, error) 반환. 예외를 밖으로 던지지 않는다."""
    try:
        r = requests.get(url, headers={"User-Agent": UA},
                         timeout=timeout, allow_redirects=True)
        return r, None
    except requests.exceptions.SSLError as e:
        return None, f"SSL 오류: {e}"
    except requests.exceptions.ConnectTimeout:
        return None, "연결 시간초과 (방화벽 또는 차단 가능성)"
    except requests.exceptions.ReadTimeout:
        return None, "응답 시간초과"
    except requests.exceptions.ConnectionError as e:
        return None, f"연결 실패: {e}"
    except requests.RequestException as e:
        return None, f"요청 오류: {e}"


def decode(resp):
    """한글 인코딩을 최대한 살려서 텍스트를 얻는다."""
    if resp.encoding and resp.encoding.lower() not in ("iso-8859-1",):
        try:
            return resp.text, resp.encoding
        except (UnicodeDecodeError, LookupError):
            pass
    for enc in ("utf-8", "euc-kr", "cp949"):
        try:
            return resp.content.decode(enc), enc
        except (UnicodeDecodeError, LookupError):
            continue
    return resp.content.decode("utf-8", errors="replace"), "utf-8(오류무시)"


def check_robots(out):
    """robots.txt 를 읽고 크롤링 허용 여부를 판단한다."""
    out("\n" + "=" * 66)
    out("  1. robots.txt 확인")
    out("=" * 66)

    for host in ("https://www.scourt.go.kr", "https://slb.scourt.go.kr"):
        url = f"{host}/robots.txt"
        out(f"\n  ▸ {url}")
        resp, err = fetch(url, timeout=15)
        if err:
            out(f"    ❌ {err}")
            continue
        if resp.status_code != 200:
            out(f"    HTTP {resp.status_code} — robots.txt 없음 (일반적으로 제한 없음)")
            continue

        text, _ = decode(resp)
        body = text.strip()
        out(f"    HTTP 200 / {len(body)} bytes")
        out("    ┌" + "─" * 56)
        for line in body.splitlines()[:25]:
            out(f"    │ {line}")
        if len(body.splitlines()) > 25:
            out(f"    │ … (총 {len(body.splitlines())}줄)")
        out("    └" + "─" * 56)

        lowered = body.lower()
        if re.search(r"disallow:\s*/\s*$", lowered, re.M):
            out("    🔴 판정: 전체 경로 크롤링 금지 — 자동수집 불가")
        elif "disallow:" in lowered:
            out("    🟡 판정: 일부 경로 금지 — 위 목록에서 대상 경로 확인 필요")
        else:
            out("    🟢 판정: 명시적 금지 없음")
        time.sleep(0.5)


def analyze(name, url, out, save_html=False):
    """대상 1건을 조사하고 결과를 출력한다. (성공 여부 반환)"""
    out(f"\n  ▸ {name}")
    out(f"    {url}")

    resp, err = fetch(url)
    if err:
        out(f"    ❌ {err}")
        return False

    final = resp.url
    if final.rstrip("/") != url.rstrip("/"):
        out(f"    ↪ 리다이렉트 → {final}")

    text, enc = decode(resp)
    size = len(resp.content)
    out(f"    HTTP {resp.status_code} / {size:,} bytes / 인코딩 {enc}")

    if resp.status_code == 404:
        out("    ⚪ 주소 없음 — 후보에서 제외")
        return False
    if resp.status_code == 403:
        out("    🔴 접근 거부 — 봇 차단 가능성. User-Agent/세션 확인 필요")
        return False
    if resp.status_code != 200:
        out(f"    🟡 예상 밖 응답 코드")
        return False

    # 사건번호가 실제로 보이는가 — 가장 결정적인 신호
    cases = CASE_RE.findall(text)
    uniq_cases = sorted(set(m.strip() for m in CASE_RE.finditer(text)
                            for m in [m.group()]))

    # 표 구조
    tables = len(re.findall(r"<table", text, re.I))
    rows = len(re.findall(r"<tr[\s>]", text, re.I))
    lists = len(re.findall(r"<ul[\s>]|<ol[\s>]", text, re.I))

    # JS 렌더링 흔적
    js_found = [h for h in JS_HINTS if h.lower() in text.lower()]

    # 본문 텍스트량 (태그 제거 후)
    plain = re.sub(r"<script.*?</script>|<style.*?</style>", " ", text,
                   flags=re.S | re.I)
    plain = re.sub(r"<[^>]+>", " ", plain)
    plain = re.sub(r"\s+", " ", plain).strip()

    out(f"    구조: <table> {tables}개 / <tr> {rows}개 / 목록 {lists}개")
    out(f"    본문 텍스트: {len(plain):,}자")
    if js_found:
        out(f"    JS 프레임워크 흔적: {', '.join(js_found)}")

    if uniq_cases:
        out(f"    🟢 사건번호 {len(uniq_cases)}건 발견 → HTML에서 직접 파싱 가능")
        out(f"       예시: {', '.join(uniq_cases[:6])}")
        verdict = "PARSEABLE"
    elif rows > 5 and len(plain) > 2000:
        out("    🟡 표 구조는 있으나 사건번호 미발견")
        out("       → 목록 페이지가 아니거나, 조회 파라미터가 필요할 수 있음")
        verdict = "MAYBE"
    elif len(plain) < 1000:
        out("    🔴 본문이 거의 비어 있음 → JS 렌더링 필요 (Selenium 대상)")
        verdict = "JS_REQUIRED"
    else:
        out("    🟡 내용은 있으나 목록 데이터로 보이지 않음")
        verdict = "MAYBE"

    # 본문 미리보기
    out("    ┌ 본문 앞부분 " + "─" * 43)
    preview = plain[:400]
    for i in range(0, len(preview), 56):
        out(f"    │ {preview[i:i + 56]}")
    out("    └" + "─" * 56)

    if save_html:
        slug = re.sub(r"[^a-zA-Z0-9]+", "_", urlparse(url).path).strip("_") or "index"
        path = OUT_DIR / "html" / f"{slug}.html"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
        out(f"    💾 저장: {path.relative_to(HERE)}")

    return verdict == "PARSEABLE"


def main():
    p = argparse.ArgumentParser(description="법원 사이트 정찰 도구")
    p.add_argument("--url", help="이 주소만 확인")
    p.add_argument("--save-html", action="store_true",
                   help="원본 HTML 저장 (파서 개발용)")
    p.add_argument("--skip-robots", action="store_true")
    args = p.parse_args()

    out = Tee(REPORT)
    try:
        out("=" * 66)
        out("  트랙 B — 법원 사이트 정찰 리포트")
        out("=" * 66)
        out(f"  실행시각: {datetime.now():%Y-%m-%d %H:%M:%S}")
        try:
            ip = requests.get("https://api.ipify.org", timeout=10).text.strip()
            out(f"  이 서버 IP: {ip}")
        except requests.RequestException:
            out("  이 서버 IP: 확인 실패 (외부 접속이 막혀 있을 수 있음)")

        if args.url:
            out("\n" + "=" * 66)
            out("  지정 주소 확인")
            out("=" * 66)
            analyze("사용자 지정", args.url, out, args.save_html)
        else:
            if not args.skip_robots:
                check_robots(out)

            out("\n" + "=" * 66)
            out("  2. 대상 사이트 접근성 확인")
            out("=" * 66)
            ok = []
            for name, url in TARGETS:
                if analyze(name, url, out, args.save_html):
                    ok.append(name)
                time.sleep(1.0)          # 서버 부담 최소화

            out("\n" + "=" * 66)
            out("  3. 종합")
            out("=" * 66)
            if ok:
                out(f"\n  🟢 바로 파싱 가능한 대상 {len(ok)}건:")
                for n in ok:
                    out(f"     · {n}")
                out("\n  → 크롤러 작성 가능. --save-html 로 다시 돌려 HTML을 확보하세요.")
            else:
                out("\n  🔴 직접 파싱 가능한 대상 없음")
                out("     원인은 위 개별 결과를 보면 판별됩니다:")
                out("       · '본문이 거의 비어 있음'  → Selenium 필요 (+2~4일)")
                out("       · 'HTTP 403'             → 봇 차단, 헤더/세션 조정 필요")
                out("       · '연결 실패/시간초과'     → VPS 방화벽 또는 사이트 차단")
                out("       · 'HTTP 404'             → 주소가 바뀜, 실제 주소 재확인 필요")

        out("\n" + "=" * 66)
        out(f"  리포트 저장: {REPORT}")
        out("  이 파일 내용을 그대로 전달하면 다음 단계로 넘어갑니다.")
        out("=" * 66)
    finally:
        out.close()

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n중단되었습니다.")
        sys.exit(130)
