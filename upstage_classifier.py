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
from compact_extraction import RULES as COMPACT_RULES, SCHEMA as COMPACT_SCHEMA, expand as expand_compact
from datetime import date, datetime
from pathlib import Path
from typing import Any

from classifier import ClassificationResult
from models import (
    CALENDAR,
    REPEAT_FREQS,
    SCOPE_REFERENCE,
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
    try:
        expanded = expand_compact(data)
    except ValueError as exc:
        raise UpstageAPIError(f"Solar Pro 4 JSON 형식 오류: {exc}") from exc
    return _to_classification(expanded, original_text)


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
                "content": COMPACT_RULES,
            },
            {"role": "user", "content": (
                f"사용자: {load_profile()}\n"
                f"기준일: {reference_date or datetime.now().astimezone().date().isoformat()}\n"
                f"발송일: {sent_at or '미상; 본문 발송일이 있으면 우선'}\n"
                f"메시지:\n{text}"
            )},
        ],
        "response_format": {"type": "json_schema", "json_schema": {
            "name": "compact_work_items", "strict": False, "schema": COMPACT_SCHEMA,
        }},
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
        try:
            try:
                response = client.chat.completions.create(**request_args)
            except Exception as exc:
                raise UpstageAPIError(f"Upstage API 호출 실패: {exc}") from exc
            return _parse_api_response(response, text)
        except UpstageAPIError as exc:
            last_error = exc
            if attempt + 1 >= MAX_API_ATTEMPTS or not _is_retryable_error(exc):
                raise
            time.sleep(0.25)
        finally:
            client.close()

    raise last_error or UpstageAPIError("Solar Pro 4 호출에 실패했습니다.")
