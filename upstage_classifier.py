"""Upstage Solar Pro 4 classifier with structured JSON output.

분류 체계
  1차 scope  : 담임 / 업무 / 참조   ← LLM 판단
  주 행동    : intent/applicability  ← 의미 보존용
  등록 대상  : type + at/due        ← 코드에서 일관성 검증

LLM에게 2차 라벨까지 물었을 때 실측 7%가 자기모순(`todo_with_due_date` 인데
due=None)이었다. 사실(at/due)만 받아 코드로 파생하면 그 모순이 불가능해진다.
"""

from __future__ import annotations

import json
import os
import re
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any

from classifier import ClassificationResult
from models import (
    CALENDAR,
    REPEAT_FREQS,
    SCOPE_REFERENCE,
    SCOPE_WORK,
    SCOPES,
    TODO,
    APPLICABILITIES,
    INTENTS,
    SOURCE_STATES,
    TEMPORAL_PRECISIONS,
    TEMPORAL_RESOLUTIONS,
    TEMPORAL_ROLES,
    Evidence,
    ReviewIssue,
    TemporalContext,
    WorkItem,
)

DEFAULT_MODEL = "solar-pro4"
UPSTAGE_BASE_URL = "https://api.upstage.ai/v1"
API_KEY_ENV_NAMES = ("UPSTAGE_API", "UPSTAGE_API_KEY", "SOLAR_API_KEY")

# solar-pro4 호출은 백그라운드 대기열에서 처리하고, 결과창은 응답이 끝난
# 뒤에만 보여 준다. 빠른 결과를 위해 추론 확장은 사용하지 않는다.
REASONING_EFFORT = "none"
# The measured p95 response time was about 50 seconds.  A 20-second cutoff
# caused otherwise valid AI results to fall back to the weaker rule parser.
REQUEST_TIMEOUT_SECONDS = 60.0
RETRY_TIMEOUT_SECONDS = 90.0
MAX_API_ATTEMPTS = 2

PROFILE_PATH = Path(__file__).resolve().parent / "profile.json"
DEFAULT_ROLE = "중학교 담임 교사"


class UpstageConfigurationError(RuntimeError):
    """Raised when the local Upstage configuration is incomplete."""


class UpstageAPIError(RuntimeError):
    """Raised when Upstage cannot produce a valid classification."""


WORK_ITEM_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "items": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "scope": {"type": "string", "enum": list(SCOPES)},
                    "type": {"type": "string", "enum": [CALENDAR, TODO]},
                    "title": {"type": "string"},
                    "detail": {"type": "string"},
                    "intent": {"type": "string", "enum": list(INTENTS)},
                    "applicability": {
                        "type": "string",
                        "enum": list(APPLICABILITIES),
                    },
                    "condition": {"type": ["string", "null"]},
                    "source_state": {
                        "type": "string",
                        "enum": list(SOURCE_STATES),
                    },
                    "at": {"type": ["string", "null"]},
                    "due": {"type": ["string", "null"]},
                    "due_time": {"type": ["string", "null"]},
                    "all_day": {"type": "boolean"},
                    "location": {"type": ["string", "null"]},
                    "reminder_minutes": {"type": ["integer", "null"]},
                    "repeat_freq": {"type": "string", "enum": list(REPEAT_FREQS)},
                    "repeat_detail": {"type": "string"},
                    "temporal_context": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "role": {"type": "string", "enum": list(TEMPORAL_ROLES)},
                                "label": {"type": "string"},
                                "raw_text": {"type": "string"},
                                "date": {"type": ["string", "null"]},
                                "time": {"type": ["string", "null"]},
                                "end_date": {"type": ["string", "null"]},
                                "end_time": {"type": ["string", "null"]},
                                "precision": {
                                    "type": "string",
                                    "enum": list(TEMPORAL_PRECISIONS),
                                },
                                "resolution": {
                                    "type": "string",
                                    "enum": list(TEMPORAL_RESOLUTIONS),
                                },
                                "segment_id": {"type": "string"},
                            },
                            "required": [
                                "role", "label", "raw_text", "date", "time",
                                "end_date", "end_time", "precision", "resolution",
                                "segment_id",
                            ],
                        },
                    },
                    "checklist": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                    "evidence": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "segment_id": {"type": "string"},
                                "field": {"type": "string"},
                                "quote": {"type": "string"},
                            },
                            "required": ["segment_id", "field", "quote"],
                        },
                    },
                    "review_issues": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "code": {"type": "string"},
                                "field": {"type": "string"},
                                "message": {"type": "string"},
                                "blocking": {"type": "boolean"},
                            },
                            "required": ["code", "field", "message", "blocking"],
                        },
                    },
                    "group_id": {"type": "string"},
                    "urgency": {"type": "string", "enum": ["즉시", "보통"]},
                },
                "required": [
                    "scope", "type", "title", "detail", "intent",
                    "applicability", "condition", "source_state", "at", "due",
                    "due_time", "all_day", "location", "reminder_minutes", "repeat_freq",
                    "repeat_detail", "temporal_context", "checklist", "evidence",
                    "review_issues", "group_id", "urgency",
                ],
            },
        }
    },
    "required": ["items"],
}

REQUIRED_ITEM_FIELDS = (
    "scope",
    "type",
    "title",
    "detail",
    "intent",
    "applicability",
    "condition",
    "source_state",
    "at",
    "due",
    "due_time",
    "all_day",
    "location",
    "reminder_minutes",
    "repeat_freq",
    "repeat_detail",
    "temporal_context",
    "checklist",
    "evidence",
    "review_issues",
    "group_id",
    "urgency",
)


def load_local_env(path: str | Path | None = None) -> None:
    """Load simple KEY=value entries without overwriting real environment vars."""

    env_path = Path(path) if path else Path(__file__).resolve().parent / ".env"
    if not env_path.is_file():
        return
    try:
        lines = env_path.read_text(encoding="utf-8").splitlines()
    except OSError:
        return
    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if stripped.startswith("export "):
            stripped = stripped[7:].lstrip()
        if "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and value:
            os.environ.setdefault(key, value)


def load_profile() -> str:
    """사용자 역할 설명. 담임/업무 판정에 필수라 없으면 정확도가 크게 떨어진다."""

    try:
        data = json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return DEFAULT_ROLE
    parts = [str(data.get("role") or DEFAULT_ROLE)]
    for key, label in (("department", "소속"), ("subject", "담당교과"), ("notes", "비고")):
        value = str(data.get(key) or "").strip()
        if value:
            parts.append(f"{label}: {value}")
    return " / ".join(parts)


def _api_key() -> str:
    load_local_env()
    for env_name in API_KEY_ENV_NAMES:
        key = os.environ.get(env_name, "").strip()
        if key:
            return key
    raise UpstageConfigurationError(
        "Upstage API 환경변수가 없습니다. UPSTAGE_API, UPSTAGE_API_KEY 또는 SOLAR_API_KEY를 설정하세요."
    )


def has_api_key() -> bool:
    """Return True when an Upstage key is available in the environment/.env."""

    try:
        _api_key()
    except UpstageConfigurationError:
        return False
    return True


def get_api_key() -> str:
    """Return the configured Upstage key for other local request pipelines."""

    return _api_key()


TAXONOMY = """## 1차 분류 (scope)
- "담임": 최종 대상이 **학생**인 일. 학생들에게 무엇을 전달·안내·배부·지도하라고 하거나,
  학생에게 전달할 사항이 있는 경우.
  (예: 가정통신문 배부, 학생들에게 설문 안내, 조회 때 전달, 학생 대상 행사 안내)
- "업무": 최종 대상이 **나(교사)** 인 일. 선생님에게 무엇을 하라거나 제출하라고 하는 경우.
  (예: 서류 제출, 연수 이수, 회의·연수 참석, 시트 입력, 신청·회신, 임장·감독)
- "참조": 읽고 알아두기만 하면 되고 내가 할 행동이 하나도 남지 않는 안내·공지

### 판정 순서
1. 학생에게 무언가를 전달·지도해야 하는가? → 담임
2. 내가 직접 하거나 제출·참석해야 하는가? → 업무
3. 둘 다 아니면 → 참조
※ 한 메시지에 학생 전달분과 내 제출분이 함께 있으면 항목을 나누어 각각 담임/업무로 넣으세요.

### 참조 판정 주의 (가장 자주 틀림)
- "학생들에게 안내/배부/지도 부탁드립니다" → 담임. 참조 아님
- "희망자는 회신 주세요", "신청 바랍니다" → 내가 대상이 될 수 있으면 참조 아님
- 내가 참석·임장해야 하는 시각이 적혀 있으면 참조 아님
- **애매하면 참조로 버리지 말고 항목으로 뽑으세요.** 불필요한 항목은 사용자가 지우면 되지만
  빠뜨린 항목은 사용자가 알아채지 못합니다. 재현율을 정확도보다 우선하세요."""

SEMANTIC_RULES = """## 행동 단위와 날짜의 관계 (가장 중요)
- 항목 수는 날짜 수·번호 수·문장 수가 아니라 **독립된 사용자 행동 수**로 정합니다.
  한 항목을 완료해도 별도의 행동이 남는 경우에만 나눕니다.
- 모집·선착순·희망자·신청·회신 요청은 `apply` 또는 `reply` 1건으로 만들고,
  모집 대상 행사일은 그 행동의 `event_context`로만 남깁니다. 행사 참석이 확정되었다는
  문장이 없으면 별도 Calendar를 만들지 않습니다.
- "반드시", "제출하세요", "기한 내", "참석 바랍니다"처럼 수신자가 해야 한다고
  명시한 행동은 `applicability=required`입니다. "희망자", "가능한 분", "해당자",
  "있으시면", "선택"처럼 조건이 붙은 행동만 `conditional`입니다.
- 모집 공지에 함께 적힌 행사일은 신청 행동의 `event_context`입니다. 행사 참석이
  확정되었다는 근거가 없으면 별도의 Calendar 항목으로 만들지 않습니다.
- 행사일과 접수 마감이 함께 있어도 같은 신청 행동에 연결합니다. `당일`, `그날`처럼
  날짜를 확정할 수 없는 마감은 `due=null`, `due_time`만 보존하고 review_issues에 확인 사유를 남깁니다.
- `type=calendar`는 사용자가 그 시각에 참석·수행하는 주 행동, `type=todo`는 신청·회신·제출·준비·안내 등
  그 전까지 처리하는 행동입니다. 관련 날짜만 있는 경우 type을 calendar로 만들지 않습니다.
- 시계 시각이 없는 날짜·기간·요일·교시·아침/오후만으로 `at`을 만들지 않습니다.
  Calendar는 원문에 정확한 시각이 있거나, 원문이 종일 행사라고 명시한 경우에만 만듭니다.
  `4교시`, `오전`, `당일`은 temporal_context에 보존하고 HH:MM으로 추정하지 않습니다.
- `at`은 주 행동의 execution, `due`와 `due_time`은 주 행동의 deadline입니다.
  나머지 모든 날짜는 temporal_context에 역할을 붙여 남깁니다:
  execution / deadline / event_context / external_deadline / constraint / historical.
- 같은 업무의 세부 입력값은 checklist로 묶습니다. 별도 산출물·별도 담당·별도 완료가 명시될 때만 분리합니다.
- 완료·전달됨·처리함만 있는 메시지는 `items=[]`로 둡니다. 아직 해야 할 후속 행동이
  함께 있으면 그 후속 행동만 `pending`으로 만들고, 이미 끝난 행동은 만들지 않습니다.
- 모든 항목은 제목·대상·시점·조건을 뒷받침하는 짧은 원문 구절을 evidence에 넣습니다.
  근거가 없거나 필수 정보가 없으면 임의로 채우지 말고 review_issues에 기록합니다.

### 대표 예시
`토익 시험 감독관 모집 / 일시 2026-08-23 08:30 / 당일 오전 10시까지 접수`
→ `토익 시험 감독관 신청` 1건, `type=todo`, `intent=apply`, `applicability=conditional`.
시험일은 event_context, 당일 10시는 unresolved deadline 문맥입니다. `접수 마감 확인`이나 시험 Calendar를 추가하지 않습니다."""

FIELD_RULES = """## 시점 — 라벨을 붙이지 말고 사실만 채우세요
- type: 주 행동이 실제로 수행되는 시점이면 calendar, 기한 전 처리하는 행동이면 todo
- at: 그 시각에 참석하거나 바로 그때 하면 되는 주 행동의 일시 "YYYY-MM-DDTHH:MM:SS".
  시각이 없는 날짜만 있으면 null로 두고 해당 날짜를 temporal_context에 남깁니다.
- due: 그때까지 끝내야 하는 주 행동의 날짜 "YYYY-MM-DD"
- due_time: 주 행동의 마감 시각 "HH:MM". 날짜가 불명확해도 시각이 있으면 보존하세요.
- at과 due가 원문에 함께 나와도 주 행동에 해당하지 않는 날짜는 temporal_context로 옮기고,
  주 행동의 type에 맞는 기본 필드만 채우세요.
- 마감이 명시되지 않았으면 due 를 지어내지 말고 null 로 두세요.
- 기준일시·발송일시는 상대 날짜 계산에만 사용합니다. 원문에 없는 날짜·시각을
  현재 날짜나 행사 날짜로 추정해 due/at에 넣지 마세요.
- `all_day=true`는 원문에 종일·하루 행사라고 명시된 Calendar에만 사용합니다.
  날짜만 있다는 이유로 `all_day=true` 또는 `00:00`을 만들지 마세요.

## 의미 필드
- intent: apply/reply/attend/submit/prepare/inform/supervise/complete_training/other 중 하나
- applicability: 항상 해야 하면 required, 희망·조건부이면 conditional, 대상이 불명확하면 unknown
- condition: 조건부 근거를 짧게. 없으면 null
- source_state: 아직 해야 하면 pending, 완료 근거면 completed, 행동 없는 공지면 informational, 판단 불가면 unknown
- review_issues: 날짜·대상·근거가 모호한 이유. `blocking=true`는 사용자가 수정하기 전 등록하지 못하는 경우에만 사용
- group_id: 같은 독립 행동의 하위 정보가 묶인 입력 내 식별자. 번호별로 무조건 새 ID를 만들지 마세요.

## 반복
- repeat_freq: 정기적으로 되풀이되는 일이면 daily / weekly / monthly, 1회성이면 none
- repeat_detail: 반복 조건을 짧게 ("화,목" / "매월 1일"). 반복이 아니면 빈 문자열
  (예: "매주 화,목요일은 분리수거날입니다" → weekly / "화,목")

## 그 밖의 필드
- title: 짧은 실행형 제목
- detail: 처리에 필요한 정보만. 인사말·발신자 소개는 넣지 마세요
- location: 장소가 있으면 넣고 없으면 null
- reminder_minutes: 몇 분 전 알림이 필요하다고 적혀 있으면 그 값, 없으면 null
- urgency: 바로/최대한 빨리 해야 하면 "즉시", 아니면 "보통"
- 한 메시지에 여러 **독립 행동**이 있으면 각각 별도 항목으로 분리하세요. 날짜·조건·세부값만 다르면 먼저 한 항목으로 묶으세요.
- JSON의 모든 필드를 빠짐없이 출력하세요. `items`가 비어 있는 경우는 행동이 없을 때만 허용합니다.
JSON 이외의 설명, Markdown 코드 펜스, 추론 내용을 출력하지 마세요."""


def _prompt(
    text: str,
    reference_date: str | None = None,
    sent_at: str | None = None,
) -> str:
    if sent_at:
        base = (
            f"[메시지 발송일시] {sent_at}\n"
            "  오늘/내일/다음주 같은 상대 날짜는 이 메시지의 발송일을 기준으로 계산하세요."
        )
        if reference_date:
            base += f"\n[사용자 지정 기준일시] {reference_date}"
    elif reference_date:
        base = (
            f"[기준일시] {reference_date}\n"
            "  이 시각을 기준으로 오늘/내일/다음주 등을 계산하세요.\n"
            "  단, 본문이나 제목에 발송일·작성일이 적혀 있으면 그 날짜를 우선하세요."
        )
    else:
        today = datetime.now().astimezone().date().isoformat()
        base = (
            f"[기준일시] {today} (오늘. 실제 발송일은 알 수 없습니다)\n"
            "  본문이나 제목에 발송일·작성일이 적혀 있으면 **그 날짜를 기준일로 삼아**\n"
            "  오늘/내일/다음주 등을 계산하세요. 그런 표기가 없을 때만 위 날짜를 쓰세요."
        )
    return (
        f"""학교 교직원 메시지에서 '내가 해야 할 일'을 뽑아 분류하세요.

[사용자] 역할: {load_profile()}
{base}

[메시지]
---
{text}
---

{TAXONOMY}

{SEMANTIC_RULES}

{FIELD_RULES}"""
    )


def _strip_json_fence(content: str) -> str:
    value = content.strip()
    if value.startswith("```"):
        value = re.sub(r"^```(?:json)?\s*", "", value, flags=re.IGNORECASE)
        value = re.sub(r"\s*```$", "", value)
    return value.strip()


def _optional_text(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _normalise_date(value: Any) -> str | None:
    text = _optional_text(value)
    if not text:
        return None
    try:
        return date.fromisoformat(text[:10]).isoformat()
    except ValueError:
        return None


def _normalise_datetime(value: Any) -> str | None:
    text = _optional_text(value)
    if not text:
        return None
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone().replace(tzinfo=None)
    return parsed.isoformat(timespec="seconds")


def _has_explicit_clock(value: Any) -> bool:
    """Return whether a datetime string contains an actual clock time."""

    text = _optional_text(value)
    return bool(text and re.search(r"(?:T|\s)\d{1,2}:\d{2}", text))


def _clock_is_in_source(value: Any, source: str) -> bool:
    """Reject a model-generated clock that never appears in the message."""

    text = _optional_text(value)
    if not text:
        return False
    match = re.search(r"(?:T|\s)(\d{1,2}):(\d{2})", text)
    if not match:
        return False
    hour, minute = int(match.group(1)), int(match.group(2))
    compact = re.sub(r"\s+", "", source)
    exact = (
        f"{hour:02d}:{minute:02d}",
        f"{hour}:{minute:02d}",
        f"{hour:02d}시{minute:02d}분",
        f"{hour}시{minute:02d}분",
    )
    if any(token in compact for token in exact):
        return True
    hour12 = hour % 12 or 12
    period_tokens = (
        f"오전{hour12}:{minute:02d}",
        f"오후{hour12}:{minute:02d}",
        f"오전{hour12}시{minute:02d}분",
        f"오후{hour12}시{minute:02d}분",
    )
    if any(token in compact for token in period_tokens):
        return True
    if minute == 0:
        if any(token in compact for token in (f"{hour:02d}시", f"{hour}시")):
            return True
        for period in ("오전", "오후"):
            if f"{period}{hour12}시" in compact:
                return True
    return False


def _normalise_reminder_minutes(value: Any) -> int | None:
    if value is None or value == "":
        return None
    try:
        minutes = int(value)
    except (TypeError, ValueError):
        return None
    return minutes if 0 < minutes <= 10080 else None


def _normalise_time(value: Any) -> str | None:
    text = _optional_text(value)
    if not text:
        return None
    match = re.fullmatch(r"(\d{1,2}):(\d{2})", text)
    if not match:
        return None
    hour, minute = int(match.group(1)), int(match.group(2))
    if hour > 23 or minute > 59:
        return None
    return f"{hour:02d}:{minute:02d}"


def _normalise_temporal_contexts(value: Any) -> list[TemporalContext]:
    if not isinstance(value, list):
        return []
    contexts: list[TemporalContext] = []
    for raw in value:
        if not isinstance(raw, dict):
            continue
        role = str(raw.get("role") or "").strip()
        if role not in TEMPORAL_ROLES:
            continue
        precision = str(raw.get("precision") or "unknown").strip()
        if precision not in TEMPORAL_PRECISIONS:
            precision = "unknown"
        resolution = str(raw.get("resolution") or "unresolved").strip()
        if resolution not in TEMPORAL_RESOLUTIONS:
            resolution = "unresolved"
        context = TemporalContext(
            role=role,
            label=str(raw.get("label") or "").strip()[:120],
            raw_text=str(raw.get("raw_text") or "").strip()[:500],
            date=_normalise_date(raw.get("date")),
            time=_normalise_time(raw.get("time")),
            end_date=_normalise_date(raw.get("end_date")),
            end_time=_normalise_time(raw.get("end_time")),
            precision=precision,
            resolution=resolution,
            segment_id=str(raw.get("segment_id") or "current").strip()[:80] or "current",
        )
        contexts.append(context)
    return contexts[:30]


def _normalise_evidence(value: Any) -> list[Evidence]:
    if not isinstance(value, list):
        return []
    evidence: list[Evidence] = []
    for raw in value:
        item = Evidence.from_dict(raw)
        if item is not None:
            evidence.append(item)
    return evidence[:30]


def _normalise_review_issues(value: Any) -> list[ReviewIssue]:
    if not isinstance(value, list):
        return []
    issues: list[ReviewIssue] = []
    for raw in value:
        item = ReviewIssue.from_dict(raw)
        if item is not None:
            issues.append(item)
    return issues[:30]


def _quote_exists(quote: str, source: str) -> bool:
    if not quote:
        return False
    if quote in source:
        return True
    compact_quote = " ".join(quote.split())
    compact_source = " ".join(source.split())
    return compact_quote in compact_source


def _to_classification(data: Any, original_text: str) -> ClassificationResult:
    if isinstance(data, list):
        data = {"items": data}
    if not isinstance(data, dict):
        raise UpstageAPIError("Solar Pro 4 응답이 JSON object가 아닙니다.")
    if "items" not in data:
        raise UpstageAPIError("Solar Pro 4 JSON에 items 필드가 없습니다.")
    raw_items = data.get("items", [])
    if not isinstance(raw_items, list):
        raise UpstageAPIError("Solar Pro 4 JSON의 items가 배열이 아닙니다.")

    work_items: list[WorkItem] = []
    ignored: list[str] = []
    for item_number, raw in enumerate(raw_items, start=1):
        if not isinstance(raw, dict):
            raise UpstageAPIError(f"Solar Pro 4 items[{item_number}]가 object가 아닙니다.")
        missing = [field for field in REQUIRED_ITEM_FIELDS if field not in raw]
        if missing:
            fields = ", ".join(missing[:5])
            suffix = " 외" if len(missing) > 5 else ""
            raise UpstageAPIError(
                f"Solar Pro 4 items[{item_number}]에 필수 필드가 없습니다: {fields}{suffix}"
            )
        title = str(raw.get("title", "")).strip()
        if not title:
            raise UpstageAPIError(f"Solar Pro 4 items[{item_number}]의 title이 비어 있습니다.")
        scope = str(raw.get("scope", "")).strip()
        if scope not in SCOPES:
            raise UpstageAPIError(f"Solar Pro 4 items[{item_number}]의 scope가 올바르지 않습니다.")

        raw_type = str(raw.get("type", "")).strip().lower()
        if raw_type not in {CALENDAR, TODO}:
            raise UpstageAPIError(f"Solar Pro 4 items[{item_number}]의 type이 올바르지 않습니다.")
        intent = str(raw.get("intent", "other")).strip()
        if intent not in INTENTS:
            raise UpstageAPIError(f"Solar Pro 4 items[{item_number}]의 intent가 올바르지 않습니다.")
        applicability = str(raw.get("applicability", "required")).strip()
        if applicability not in APPLICABILITIES:
            raise UpstageAPIError(
                f"Solar Pro 4 items[{item_number}]의 applicability가 올바르지 않습니다."
            )
        source_state = str(raw.get("source_state", "pending")).strip()
        if source_state not in SOURCE_STATES:
            raise UpstageAPIError(
                f"Solar Pro 4 items[{item_number}]의 source_state가 올바르지 않습니다."
            )

        all_day = raw.get("all_day")
        if not isinstance(all_day, bool):
            raise UpstageAPIError(f"Solar Pro 4 items[{item_number}]의 all_day가 boolean이 아닙니다.")

        raw_at = _optional_text(raw.get("at"))
        raw_due = _optional_text(raw.get("due"))
        raw_due_time = _optional_text(raw.get("due_time"))
        condition = _optional_text(raw.get("condition"))
        at = _normalise_datetime(raw_at)
        due = _normalise_date(raw_due)
        due_time = _normalise_time(raw_due_time)
        contexts = _normalise_temporal_contexts(raw.get("temporal_context"))
        issues = _normalise_review_issues(raw.get("review_issues"))
        evidence = _normalise_evidence(raw.get("evidence"))
        derived_issues: list[ReviewIssue] = list(issues)

        uncertainty_codes = {
            "missing_execution_time",
            "relative_date",
            "ambiguous_deadline_date",
            "uncertain_time",
        }
        for issue in derived_issues:
            message = issue.message.casefold()
            if (
                issue.code.casefold() in uncertainty_codes
                or "추정" in message
                or "불확실" in message
                or "시각이 없어" in message
            ):
                # The model has already admitted that this field is uncertain.
                # Preserve the item for editing, but prevent one-click saving.
                issue.blocking = True

        def add_issue(issue: ReviewIssue) -> None:
            if not any(existing.code == issue.code and existing.field == issue.field for existing in derived_issues):
                derived_issues.append(issue)

        if raw_at and at is None:
            add_issue(ReviewIssue(
                code="invalid_execution_time",
                field="at",
                message="수행 시각 형식을 이해하지 못해 등록할 수 없습니다.",
                blocking=True,
            ))
        if raw_due and due is None:
            add_issue(ReviewIssue(
                code="invalid_deadline_date",
                field="due",
                message="마감 날짜 형식을 이해하지 못해 등록할 수 없습니다.",
                blocking=True,
            ))
        if raw_due_time and due_time is None:
            add_issue(ReviewIssue(
                code="invalid_deadline_time",
                field="due_time",
                message="마감 시각 형식을 이해하지 못해 등록할 수 없습니다.",
                blocking=True,
            ))
        checklist_raw = raw.get("checklist", [])
        checklist = (
            [str(value).strip() for value in checklist_raw if str(value).strip()]
            if isinstance(checklist_raw, list)
            else []
        )[:30]

        # Keep the primary target separate from related dates. When the model
        # sends both at and due, the non-primary one becomes context instead
        # of generating a second WorkItem or silently losing the date.
        if raw_type == CALENDAR:
            primary_start = at
            primary_due = None
            if at and raw_at and not _has_explicit_clock(raw_at) and not all_day:
                contexts.append(
                    TemporalContext(
                        role="execution",
                        label="시각 없는 수행 날짜",
                        raw_text=raw_at,
                        date=at[:10],
                        precision="date",
                        resolution="explicit",
                    )
                )
                primary_start = None
                add_issue(ReviewIssue(
                    code="missing_execution_time",
                    field="start",
                    message="날짜만 있어 Calendar 시각을 정할 수 없습니다. 시각을 확인하세요.",
                    blocking=True,
                ))
            elif (
                at
                and raw_at
                and not all_day
                and not _clock_is_in_source(raw_at, original_text)
            ):
                contexts.append(
                    TemporalContext(
                        role="execution",
                        label="원문에서 확인되지 않은 수행 시각",
                        raw_text=raw_at,
                        date=at[:10],
                        time=at[11:16],
                        precision="datetime",
                        resolution="inferred",
                    )
                )
                primary_start = None
                add_issue(ReviewIssue(
                    code="execution_time_not_in_source",
                    field="start",
                    message="모델이 만든 수행 시각을 원문에서 확인하지 못했습니다. 시각을 확인하세요.",
                    blocking=True,
                ))
            if due or due_time:
                contexts.append(
                    TemporalContext(
                        role="deadline",
                        label="관련 마감",
                        raw_text="",
                        date=due,
                        time=due_time,
                        precision="datetime" if due and due_time else "date" if due else "unknown",
                        resolution="explicit" if due or due_time else "unresolved",
                    )
                )
                due = None
                due_time = None
        else:
            primary_start = None
            primary_due = due
            if at:
                contexts.append(
                    TemporalContext(
                        role="execution",
                        label="관련 수행 시각",
                        raw_text=raw_at or "",
                        date=at[:10],
                        time=at[11:16] if raw_at and _has_explicit_clock(raw_at) else None,
                        precision="datetime" if raw_at and _has_explicit_clock(raw_at) else "date",
                        resolution="explicit",
                    )
                )
                at = None
            if all_day:
                add_issue(ReviewIssue(
                    code="all_day_not_calendar",
                    field="all_day",
                    message="종일 표시는 Calendar 항목에서만 사용할 수 있습니다.",
                    blocking=True,
                ))
                all_day = False

        if applicability == "required" and condition:
            # The model supplied a condition but labelled the item required.
            # Treat the safer interpretation as conditional until the user
            # confirms it in the card.
            applicability = "conditional"
            add_issue(ReviewIssue(
                code="required_condition_conflict",
                field="applicability",
                message="조건이 함께 제시되어 조건부 항목으로 보류했습니다.",
                blocking=False,
            ))
        if applicability == "conditional" and not condition:
            add_issue(ReviewIssue(
                code="missing_condition",
                field="condition",
                message="조건부 항목의 조건 근거가 없어 확인이 필요합니다.",
                blocking=True,
            ))

        group_id = _optional_text(raw.get("group_id"))
        item = WorkItem(
            type=raw_type,
            scope=scope,
            title=title[:240],
            start=primary_start,
            due=primary_due,
            due_time=due_time,
            description=str(raw.get("detail", "")).strip()[:4000],
            location=_optional_text(raw.get("location")),
            all_day=all_day,
            reminder_minutes=_normalise_reminder_minutes(raw.get("reminder_minutes")),
            repeat_freq=(
                str(raw.get("repeat_freq", "none")).strip()
                if str(raw.get("repeat_freq", "none")).strip() in REPEAT_FREQS
                else "none"
            ),
            repeat_detail=str(raw.get("repeat_detail", "")).strip()[:300],
            intent=intent,
            applicability=applicability,
            condition=condition,
            source_state=source_state,
            temporal_context=contexts[:30],
            checklist=checklist,
            evidence=evidence,
            review_issues=derived_issues,
            group_id=group_id,
        )
        if item.scope == SCOPE_REFERENCE:
            ignored.append(title)
            item.selected = False
        if item.applicability != "required" or item.source_state != "pending":
            item.selected = False
        if item.type == CALENDAR and not item.start:
            add_issue(ReviewIssue(
                code="missing_execution_time",
                field="start",
                message="수행·참석 시각이 없어 등록 전에 확인이 필요합니다.",
                blocking=True,
            )
            )
            item.selected = False
        for evidence_item in item.evidence:
            if not _quote_exists(evidence_item.quote, original_text):
                add_issue(ReviewIssue(
                    code="evidence_not_found",
                    field=evidence_item.field,
                    message="모델이 제시한 근거 문장을 원문에서 그대로 확인하지 못했습니다.",
                    blocking=False,
                )
                )
                break
        if not item.evidence:
            add_issue(ReviewIssue(
                code="missing_evidence",
                field="evidence",
                message="추출 근거 문장이 없어 원문을 확인해 주세요.",
                blocking=False,
            )
            )
        if item.blocking_review_issues:
            item.selected = False
        work_items.append(item)

    return ClassificationResult(work_items=work_items, ignored=ignored)


def _parse_api_response(response: Any, original_text: str) -> ClassificationResult:
    try:
        content = response.choices[0].message.content
        if not content:
            raise ValueError("응답 content가 비어 있습니다.")
        data = json.loads(_strip_json_fence(content))
    except (AttributeError, IndexError, TypeError, ValueError, json.JSONDecodeError) as exc:
        raise UpstageAPIError(f"Solar Pro 4 응답 JSON 파싱 실패: {exc}") from exc
    return _to_classification(data, original_text)


def _is_retryable_error(error: Exception) -> bool:
    message = str(error).casefold()
    retry_markers = (
        "timeout",
        "timed out",
        "time-out",
        "connection",
        "temporarily",
        "429",
        "502",
        "503",
        "504",
        "json",
        "items 필드",
        "필수 필드",
    )
    return any(marker in message for marker in retry_markers)


def classify_with_upstage(
    text: str,
    *,
    model: str | None = None,
    api_key: str | None = None,
    reference_date: str | None = None,
    sent_at: str | None = None,
) -> ClassificationResult:
    """Call Solar Pro 4 and convert its JSON response into WorkItems."""

    if not text or not text.strip():
        raise ValueError("분석할 텍스트가 비어 있습니다.")
    key = api_key.strip() if api_key else _api_key()
    selected_model = model or os.environ.get("UPSTAGE_MODEL", DEFAULT_MODEL)

    try:
        from openai import OpenAI
    except ImportError as exc:  # pragma: no cover - depends on the local environment
        raise UpstageConfigurationError(
            "openai 패키지가 없습니다. `python -m pip install -r requirements.txt`를 실행하세요."
        ) from exc

    request_args: dict[str, Any] = {
        "model": selected_model,
        "messages": [
            {
                "role": "system",
                "content": "학교 교직원 메시지를 정해진 스키마로 분류하는 도구입니다.",
            },
            {"role": "user", "content": _prompt(text, reference_date, sent_at)},
        ],
        "response_format": {
            "type": "json_schema",
            "json_schema": {
                "name": "work_item_extraction",
                "strict": True,
                "schema": WORK_ITEM_SCHEMA,
            },
        },
        "reasoning_effort": REASONING_EFFORT,
    }
    last_error: UpstageAPIError | None = None
    for attempt in range(MAX_API_ATTEMPTS):
        timeout = RETRY_TIMEOUT_SECONDS if attempt else REQUEST_TIMEOUT_SECONDS
        client = OpenAI(
            api_key=key,
            base_url=UPSTAGE_BASE_URL,
            timeout=timeout,
            max_retries=0,
        )
        attempt_args = dict(request_args)
        try:
            try:
                response = client.chat.completions.create(**attempt_args)
            except Exception as exc:
                message = str(exc)
                if "response_format" not in message.casefold() and "json_schema" not in message.casefold():
                    raise UpstageAPIError(f"Upstage API 호출 실패: {message}") from exc
                # Older OpenAI-compatible gateways may accept JSON object mode
                # but not the stricter schema wrapper.
                attempt_args["response_format"] = {"type": "json_object"}
                try:
                    response = client.chat.completions.create(**attempt_args)
                except Exception as fallback_exc:
                    raise UpstageAPIError(f"Upstage JSON 호출 실패: {fallback_exc}") from fallback_exc
            return _parse_api_response(response, text)
        except UpstageAPIError as exc:
            last_error = exc
            if attempt + 1 >= MAX_API_ATTEMPTS or not _is_retryable_error(exc):
                raise
            time.sleep(0.25)

    raise last_error or UpstageAPIError("Solar Pro 4 호출에 실패했습니다.")
