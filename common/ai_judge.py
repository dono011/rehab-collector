# -*- coding: utf-8 -*-
"""
AI 판정 모듈 — 회생·매각 공고 후보를 AI로 재판단한다.

트랙 A(DART)·트랙 B(법원공고)가 이름/업종코드 정규식으로 1차 선별한
"전기 관련 후보(관련도 2 이상)"만 대상으로, 세 제공자 중 하나에게
관련도·위험신호·판단근거를 다시 물어본다.

    jev     TypeSafe Jev — 저렴한 정형 판단(관련도/위험 분류) 전용.
            매일 자동으로 도는 판정은 이걸로 충분하고 비용이 훨씬 싸다.
    claude  Anthropic Claude — 서술형 배경조사·판례 해석이 필요할 때만.
    openai  OpenAI GPT — claude와 같은 용도의 대안.

환경변수로 고른다 (기본값 jev):

    export AI_JUDGE_PROVIDER=jev      # 또는 claude / openai
    export TYPESAFE_AI_API_KEY="..."  # jev — TypeSafe 직접 발급 키 (있으면 우선 사용)
    export OPENROUTER_API_KEY="..."   # jev — TypeSafe 신규가입 중단 시 우회 경로.
                                       #        TYPESAFE_AI_API_KEY가 없으면 이걸로 자동 전환.
    export ANTHROPIC_API_KEY="..."    # claude
    export OPENAI_API_KEY="..."       # openai

Jev는 OpenRouter에서도 서빙되지만, 구조화 판정 모델이라 OpenRouter의
일반 chat completions API로는 호출할 수 없다. TypeSafe 호환 전용
엔드포인트(openrouter.ai/api/v1/systemone, 2026-09 기준 alpha)를 그대로
쓴다 — 아직 alpha라 스펙이 바뀔 수 있다.

⛔ 모델 이름은 "jev-latest" 다 (2026-09-26 실측). "typesafe/" 를 붙이면
   HTTP 400 "Model ... does not exist" 가 난다.

이 모듈은 절대 예외를 밖으로 던지지 않는다 — 크론으로 매일 도는
수집 스크립트가 AI 판정 실패(키 미설정, 네트워크 오류, 가입 중단 등)로
멈추면 안 되기 때문이다. 실패하면 "_error" 키가 담긴 dict를 돌려주고,
호출부는 그 후보만 건너뛰면 된다.
"""

import json
import os

import requests

TYPESAFE_URL = "https://api.typesafe.ai/v1/systemone"
OPENROUTER_SYSTEMONE_URL = "https://openrouter.ai/api/v1/systemone"
ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
OPENAI_URL = "https://api.openai.com/v1/chat/completions"

TIMEOUT = 30

INSTRUCTIONS = {
    "relevance": ("이 회생/파산/매각 공고가 전기공사업(전기공사·전기설비·배전·송전·"
                  "변전·수변전) 관련 회사에 대한 것인지 판단해줘."),
    "risk": ("공고·회사 정보에 인수를 재고해야 할 위험 신호(완전자본잠식, 거액 "
             "우발채무·PF지급보증, 경영권 분쟁·소송, 등록취소 사유)가 보이는지 "
             "판단해줘."),
}

RELEVANCE_CRITERIA = ["높음", "보통", "낮음", "무관"]
RISK_CRITERIA = ["위험", "주의", "낮음"]

JSON_PROMPT_SUFFIX = (
    '\n\n위 내용을 보고 아래 JSON 형식으로만 답해줘 (다른 텍스트 없이, 코드블록 없이):\n'
    '{"관련도": "높음|보통|낮음|무관", "위험신호": "위험|주의|낮음", "판단근거": "한국어 한 문장"}'
)


def _build_state_text(company, title, context=""):
    text = f"회사명: {company}\n공고 제목: {title}"
    if context:
        text += f"\n추가 정보: {context[:1500]}"
    return text


def judge_candidate(company, title, context="", provider=None):
    """
    후보 1건을 AI로 재판정한다.

    성공 시: {"관련도": str, "위험신호": str, "판단근거": str, "_provider": str}
    실패 시: {"_error": str, "_provider": str}
    """
    provider = (provider or os.environ.get("AI_JUDGE_PROVIDER", "jev")).strip().lower()
    state_text = _build_state_text(company, title, context)

    handlers = {"jev": _judge_jev, "claude": _judge_claude, "openai": _judge_openai}
    handler = handlers.get(provider)
    if not handler:
        return {"_error": f"알 수 없는 provider: {provider} (jev/claude/openai 중 하나)",
                "_provider": provider}

    try:
        return handler(state_text)
    except Exception as e:
        return {"_error": f"{e.__class__.__name__}: {e}", "_provider": provider}


def _jev_transport():
    """
    TypeSafe 직접 키가 있으면 그걸 쓰고, 없으면 OpenRouter 키로 우회한다.

    TypeSafe 신규 가입이 막혀 있어도(2026-09-22부터 일시 중단) OpenRouter는
    대기 없이 키를 받을 수 있어 이쪽으로 자동 전환한다. 나중에 TypeSafe
    가입이 재개돼 TYPESAFE_AI_API_KEY를 등록하면 자동으로 직접 호출로
    돌아간다 — 우선순위: TypeSafe 직접 키 > OpenRouter.

    반환: (url, headers, model, provider_label) 또는 (None, None, None, 에러메시지)
    """
    ts_key = os.environ.get("TYPESAFE_AI_API_KEY", "").strip()
    if ts_key:
        return (TYPESAFE_URL,
                {"Authorization": f"Bearer {ts_key}", "Content-Type": "application/json"},
                os.environ.get("JEV_MODEL", "jev-latest"),
                "jev")

    or_key = os.environ.get("OPENROUTER_API_KEY", "").strip()
    if or_key:
        # ⛔ 모델 이름에 "typesafe/" 를 붙이면 안 된다 (2026-09-26 실측).
        #    systemone 엔드포인트는 "jev-latest" 만 받는다 —
        #    "typesafe/jev-latest", "typesafe/jev-router" 는 HTTP 400
        #    "Model ... does not exist". (typesafe/jev-router 는 OpenRouter
        #    일반 모델목록에는 있지만 systemone 쪽에서는 안 먹는다.)
        #    실제로는 typesafe/jev-1.13-20260917 로 연결된다.
        return (OPENROUTER_SYSTEMONE_URL,
                {"Authorization": f"Bearer {or_key}", "Content-Type": "application/json"},
                os.environ.get("JEV_MODEL", "jev-latest"),
                "jev(openrouter)")

    return None, None, None, "TYPESAFE_AI_API_KEY 또는 OPENROUTER_API_KEY 필요"


def _judge_jev(state_text):
    url, headers, model, provider_or_error = _jev_transport()
    if url is None:
        return {"_error": provider_or_error, "_provider": "jev"}
    provider = provider_or_error

    body = {
        "model": model,
        "state": {"document": state_text},
        "questions": {
            "relevance": {
                "type": "choice",
                "instructions": INSTRUCTIONS["relevance"],
                "criteria": {c: None for c in RELEVANCE_CRITERIA},
            },
            "risk": {
                "type": "choice",
                "instructions": INSTRUCTIONS["risk"],
                "criteria": {c: None for c in RISK_CRITERIA},
            },
        },
    }
    r = requests.post(url, json=body, timeout=TIMEOUT, headers=headers)
    if r.status_code == 401:
        return {"_error": "HTTP 401 — API 키를 확인하세요", "_provider": provider}
    if r.status_code != 200:
        return {"_error": f"HTTP {r.status_code}: {r.text[:200]}", "_provider": provider}

    data = r.json()
    answers = data.get("answers", {})
    relevance = answers.get("relevance") or {}
    risk = answers.get("risk") or {}
    return {
        "관련도": relevance.get("choice", ""),
        "위험신호": risk.get("choice", ""),
        "판단근거": _confidence_note(relevance, risk),
        "_provider": provider,
        "_usage": data.get("usage", {}),
    }


def _confidence_note(relevance, risk):
    """
    Jev는 서술형 근거를 만들지 못한다 (2026-09-26 실측).

    systemone 의 질문 유형은 noul / choice / score 세 가지뿐이고,
    noul 은 글이 아니라 0~1 숫자를 돌려준다 — "판단 근거를 한 문장으로"
    라고 물어도 `0.45` 가 온다. System One(직관 판정) 모델이라 그렇다.
    서술형 근거가 필요하면 AI_JUDGE_PROVIDER=claude|openai 를 쓴다.

    그래서 근거 자리에는 판정의 확신도를 적는다.
    """
    def pct(answer):
        c = answer.get("confidence")
        return f"{c * 100:.0f}%" if isinstance(c, (int, float)) else "?"

    return f"Jev 확신도 — 관련도 {pct(relevance)} · 위험 {pct(risk)}"


def _parse_json_reply(text, provider):
    start, end = text.find("{"), text.rfind("}")
    if start == -1 or end == -1:
        return {"_error": f"JSON 형식이 아닌 응답: {text[:200]}", "_provider": provider}
    parsed = json.loads(text[start:end + 1])
    parsed["_provider"] = provider
    return parsed


def _judge_claude(state_text):
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not key:
        return {"_error": "ANTHROPIC_API_KEY 없음", "_provider": "claude"}

    prompt = state_text + JSON_PROMPT_SUFFIX
    r = requests.post(ANTHROPIC_URL, timeout=TIMEOUT, headers={
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
        "Content-Type": "application/json",
    }, json={
        "model": os.environ.get("CLAUDE_MODEL", "claude-haiku-4-5-20251001"),
        "max_tokens": 300,
        "messages": [{"role": "user", "content": prompt}],
    })
    if r.status_code != 200:
        return {"_error": f"HTTP {r.status_code}: {r.text[:200]}", "_provider": "claude"}

    text = r.json()["content"][0]["text"]
    return _parse_json_reply(text, "claude")


def _judge_openai(state_text):
    key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not key:
        return {"_error": "OPENAI_API_KEY 없음", "_provider": "openai"}

    prompt = state_text + JSON_PROMPT_SUFFIX
    r = requests.post(OPENAI_URL, timeout=TIMEOUT, headers={
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }, json={
        "model": os.environ.get("OPENAI_MODEL", "gpt-4o-mini"),
        "response_format": {"type": "json_object"},
        "messages": [{"role": "user", "content": prompt}],
    })
    if r.status_code != 200:
        return {"_error": f"HTTP {r.status_code}: {r.text[:200]}", "_provider": "openai"}

    text = r.json()["choices"][0]["message"]["content"]
    return _parse_json_reply(text, "openai")
