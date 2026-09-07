"""Data models shared by the classifier, UI, storage, and Outlook adapter."""

from __future__ import annotations

from dataclasses import dataclass, field
import json
import re
from typing import Any


CALENDAR = "calendar"
TODO = "todo"
WORK_ITEM_TYPES = (CALENDAR, TODO)

# 1차 분류 — 누구를 위한 일인가
SCOPE_HOMEROOM = "담임"   # 학생에게 전달·안내·배부·지도해야 하는 일
SCOPE_WORK = "업무"       # 내가 직접 하거나 제출·참석해야 하는 일
SCOPE_REFERENCE = "참조"  # 읽고 넘기면 되는 안내. Outlook에 등록하지 않는다
SCOPES = (SCOPE_HOMEROOM, SCOPE_WORK, SCOPE_REFERENCE)

# 2차 분류 — 필드값에서 파생한다. 직접 입력받지 않는다.
ACTION_TODO = "todo"                    # 마감 없음 / 바로 처리
ACTION_WITH_DATE = "todo_with_date"     # 그 시각에 참석·수행
ACTION_WITH_DUE = "todo_with_due_date"  # 그때까지 마감
ACTION_NONE = "none"                    # 참조

REPEAT_FREQS = ("none", "daily", "weekly", "monthly")
OUTLOOK_STATUSES = ("pending", "saved", "failed", "deleted")

# 의미 추출 v2.  `action`은 화면/Outlook 대상용 파생값으로 남겨 두고,
# 실제로 어떤 행동인지와 적용 조건은 별도 필드에 보존한다.
INTENTS = (
    "apply",
    "reply",
    "attend",
    "submit",
    "prepare",
    "inform",
    "supervise",
    "complete_training",
    "other",
)
APPLICABILITIES = ("required", "conditional", "unknown")
SOURCE_STATES = ("pending", "completed", "informational", "unknown")
TEMPORAL_ROLES = (
    "execution",
    "deadline",
    "event_context",
    "external_deadline",
    "constraint",
    "historical",
)
TEMPORAL_PRECISIONS = ("datetime", "date", "period", "school_period", "unknown")
TEMPORAL_RESOLUTIONS = ("explicit", "inferred", "unresolved")


@dataclass
class TemporalContext:
    """A date/time mentioned in the source, including non-registration dates."""

    role: str
    label: str = ""
    raw_text: str = ""
    date: str | None = None
    time: str | None = None
    end_date: str | None = None
    end_time: str | None = None
    precision: str = "unknown"
    resolution: str = "unresolved"
    segment_id: str = "current"

    def to_dict(self) -> dict[str, Any]:
        return {
            "role": self.role,
            "label": self.label,
            "raw_text": self.raw_text,
            "date": self.date,
            "time": self.time,
            "end_date": self.end_date,
            "end_time": self.end_time,
            "precision": self.precision,
            "resolution": self.resolution,
            "segment_id": self.segment_id,
        }

    @classmethod
    def from_dict(cls, value: Any) -> "TemporalContext | None":
        if not isinstance(value, dict):
            return None
        role = str(value.get("role") or "").strip()
        if role not in TEMPORAL_ROLES:
            role = "historical"
        precision = str(value.get("precision") or "unknown").strip()
        if precision not in TEMPORAL_PRECISIONS:
            precision = "unknown"
        resolution = str(value.get("resolution") or "unresolved").strip()
        if resolution not in TEMPORAL_RESOLUTIONS:
            resolution = "unresolved"
        return cls(
            role=role,
            label=str(value.get("label") or "").strip()[:120],
            raw_text=str(value.get("raw_text") or "").strip()[:500],
            date=_optional_string(value.get("date")),
            time=_optional_string(value.get("time")),
            end_date=_optional_string(value.get("end_date")),
            end_time=_optional_string(value.get("end_time")),
            precision=precision,
            resolution=resolution,
            segment_id=str(value.get("segment_id") or "current").strip()[:80] or "current",
        )


@dataclass
class Evidence:
    """A short source quote supporting one extracted field."""

    segment_id: str = "current"
    field: str = ""
    quote: str = ""

    def to_dict(self) -> dict[str, str]:
        return {
            "segment_id": self.segment_id,
            "field": self.field,
            "quote": self.quote,
        }

    @classmethod
    def from_dict(cls, value: Any) -> "Evidence | None":
        if not isinstance(value, dict):
            return None
        quote = str(value.get("quote") or "").strip()
        if not quote:
            return None
        return cls(
            segment_id=str(value.get("segment_id") or "current").strip()[:80] or "current",
            field=str(value.get("field") or "").strip()[:80],
            quote=quote[:500],
        )


@dataclass
class ReviewIssue:
    """A machine-detected uncertainty that should be visible before saving."""

    code: str
    field: str = ""
    message: str = ""
    blocking: bool = False
    resolved: bool = False

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "field": self.field,
            "message": self.message,
            "blocking": self.blocking,
            "resolved": self.resolved,
        }

    @classmethod
    def from_dict(cls, value: Any) -> "ReviewIssue | None":
        if not isinstance(value, dict):
            return None
        code = str(value.get("code") or "unknown").strip()[:80]
        return cls(
            code=code or "unknown",
            field=str(value.get("field") or "").strip()[:80],
            message=str(value.get("message") or "").strip()[:500],
            blocking=bool(value.get("blocking", False)),
            resolved=bool(value.get("resolved", False)),
        )


def _optional_string(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _json_value(value: Any, default: Any) -> Any:
    if value is None or value == "":
        return default
    try:
        return json.loads(str(value))
    except (TypeError, ValueError, json.JSONDecodeError):
        return default


def _time_is_valid(value: str | None) -> bool:
    if not value:
        return True
    return bool(re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", value.strip()))


@dataclass
class WorkItem:
    type: str
    title: str
    scope: str = SCOPE_WORK
    start: str | None = None
    end: str | None = None
    due: str | None = None
    description: str = ""
    selected: bool = True
    saved: bool = False
    id: int | None = None
    input_id: int | None = None
    outlook_entry_id: str | None = None
    last_error: str | None = None
    outlook_status: str = "pending"
    # Outlook appointment/task metadata. These fields are optional so older
    # JSON and existing callers remain compatible.
    location: str | None = None
    category: str | None = None
    all_day: bool = False
    reminder_minutes: int | None = None
    reminder: str | None = None
    repeat_freq: str = "none"
    repeat_detail: str = ""
    # UI-only guard.  A draft can be excluded before it reaches Outlook when
    # an existing item occupies the same slot or is an exact Todo duplicate.
    excluded_reason: str | None = None
    # Semantic extraction v2. These defaults keep older rows and JSON input
    # compatible while preserving the meaning behind a registration draft.
    intent: str = "other"
    applicability: str = "required"
    condition: str | None = None
    source_state: str = "pending"
    due_time: str | None = None
    temporal_context: list[TemporalContext] = field(default_factory=list)
    checklist: list[str] = field(default_factory=list)
    evidence: list[Evidence] = field(default_factory=list)
    review_issues: list[ReviewIssue] = field(default_factory=list)
    group_id: str | None = None

    @property
    def action(self) -> str:
        """2차 분류. 라벨을 저장하지 않고 필드에서 매번 파생한다.

        LLM에게 라벨까지 물으면 `todo_with_due_date` 인데 due=None 같은
        자기모순이 실측 7% 나왔다. 파생하면 그 모순이 구조적으로 불가능하다.
        """

        if self.scope == SCOPE_REFERENCE:
            return ACTION_NONE
        if self.outlook_target == CALENDAR:
            return ACTION_WITH_DATE
        if self.due:
            return ACTION_WITH_DUE
        return ACTION_TODO

    @property
    def outlook_target(self) -> str | None:
        """Outlook 어디에 넣을지. 참조는 None(등록하지 않음)."""

        if self.scope == SCOPE_REFERENCE:
            return None
        # An unresolved Calendar candidate is retained during analysis and is
        # blocked at registration until a start time is supplied.
        if self.type == CALENDAR:
            return CALENDAR
        if self.due:
            return TODO
        if self.start:
            return CALENDAR
        return TODO

    @property
    def blocking_review_issues(self) -> list[ReviewIssue]:
        return [
            issue
            for issue in self.review_issues
            if issue.blocking and not issue.resolved
        ]

    @property
    def needs_review(self) -> bool:
        return bool(self.blocking_review_issues) or self.applicability == "unknown"

    @property
    def can_register(self) -> bool:
        return (
            self.outlook_target is not None
            and self.applicability != "unknown"
            and self.source_state not in {"completed", "informational"}
            and not self.blocking_review_issues
        )

    def semantic_payload(self) -> dict[str, Any]:
        """JSON-safe v2 fields stored in SQLite's semantic_json column."""

        return {
            "intent": self.intent,
            "applicability": self.applicability,
            "condition": self.condition,
            "source_state": self.source_state,
            "temporal_context": [context.to_dict() for context in self.temporal_context],
            "checklist": list(self.checklist),
            "evidence": [item.to_dict() for item in self.evidence],
            "review_issues": [issue.to_dict() for issue in self.review_issues],
            "group_id": self.group_id,
        }

    def validate(self) -> None:
        if self.type not in WORK_ITEM_TYPES:
            raise ValueError(f"지원하지 않는 WorkItem type입니다: {self.type}")
        if self.scope not in SCOPES:
            raise ValueError(f"지원하지 않는 scope입니다: {self.scope}")
        if self.repeat_freq not in REPEAT_FREQS:
            raise ValueError(f"지원하지 않는 repeat_freq입니다: {self.repeat_freq}")
        if not self.title.strip():
            raise ValueError("제목은 비워둘 수 없습니다.")
        if self.type == CALENDAR and not self.start:
            raise ValueError("Calendar 항목의 수행 날짜와 시간이 필요합니다.")
        if self.all_day and self.type != CALENDAR:
            raise ValueError("종일 일정은 Calendar 항목에만 사용할 수 있습니다.")
        if self.reminder_minutes is not None and not 0 <= self.reminder_minutes <= 10080:
            raise ValueError("알림 시간은 0~10080분 사이여야 합니다.")
        if self.intent not in INTENTS:
            raise ValueError(f"지원하지 않는 intent입니다: {self.intent}")
        if self.applicability not in APPLICABILITIES:
            raise ValueError(f"지원하지 않는 applicability입니다: {self.applicability}")
        if self.source_state not in SOURCE_STATES:
            raise ValueError(f"지원하지 않는 source_state입니다: {self.source_state}")
        if not _time_is_valid(self.due_time):
            raise ValueError("due_time은 HH:MM 형식이어야 합니다.")

    @classmethod
    def from_row(cls, row: Any) -> "WorkItem":
        def value(name: str, default: Any = None) -> Any:
            try:
                return row[name]
            except (IndexError, KeyError):
                return default

        outlook_status = str(
            value("outlook_status")
            or ("saved" if value("saved", 0) else "pending")
        )
        semantic = _json_value(value("semantic_json"), {})
        if not isinstance(semantic, dict):
            semantic = {}
        contexts: list[TemporalContext] = []
        for raw_context in semantic.get("temporal_context", []):
            context = TemporalContext.from_dict(raw_context)
            if context is not None:
                contexts.append(context)
        evidence: list[Evidence] = []
        for raw_evidence in semantic.get("evidence", []):
            evidence_item = Evidence.from_dict(raw_evidence)
            if evidence_item is not None:
                evidence.append(evidence_item)
        issues: list[ReviewIssue] = []
        for raw_issue in semantic.get("review_issues", []):
            issue = ReviewIssue.from_dict(raw_issue)
            if issue is not None:
                issues.append(issue)
        checklist = semantic.get("checklist", [])
        if not isinstance(checklist, list):
            checklist = []
        return cls(
            id=int(value("id")),
            input_id=int(value("input_id")),
            type=str(value("type", "")),
            scope=str(value("scope") or SCOPE_WORK),
            repeat_freq=str(value("repeat_freq") or "none"),
            repeat_detail=str(value("repeat_detail") or ""),
            title=str(value("title", "")),
            start=value("start"),
            end=value("end"),
            due=value("due"),
            description=str(value("description", "") or ""),
            saved=bool(value("saved", 0)),
            outlook_entry_id=value("outlook_entry_id"),
            last_error=value("last_error"),
            outlook_status=outlook_status,
            location=value("location"),
            category=value("category"),
            all_day=bool(value("all_day", 0)),
            reminder_minutes=(
                int(value("reminder_minutes"))
                if value("reminder_minutes") is not None
                else None
            ),
            reminder=value("reminder"),
            selected=not bool(value("saved", 0)) and outlook_status != "deleted",
            intent=str(semantic.get("intent") or "other"),
            applicability=str(semantic.get("applicability") or "required"),
            condition=_optional_string(semantic.get("condition")),
            source_state=str(semantic.get("source_state") or "pending"),
            due_time=_optional_string(value("due_time")),
            temporal_context=contexts,
            checklist=[str(item).strip() for item in checklist if str(item).strip()][:30],
            evidence=evidence,
            review_issues=issues,
            group_id=_optional_string(semantic.get("group_id")),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "input_id": self.input_id,
            "type": self.type,
            "scope": self.scope,
            "action": self.action,
            "title": self.title,
            "start": self.start,
            "end": self.end,
            "due": self.due,
            "description": self.description,
            "selected": self.selected,
            "saved": self.saved,
            "outlook_entry_id": self.outlook_entry_id,
            "last_error": self.last_error,
            "outlook_status": self.outlook_status,
            "location": self.location,
            "category": self.category,
            "all_day": self.all_day,
            "reminder_minutes": self.reminder_minutes,
            "reminder": self.reminder,
            "repeat_freq": self.repeat_freq,
            "repeat_detail": self.repeat_detail,
            "excluded_reason": self.excluded_reason,
            "intent": self.intent,
            "applicability": self.applicability,
            "condition": self.condition,
            "source_state": self.source_state,
            "due_time": self.due_time,
            "temporal_context": [context.to_dict() for context in self.temporal_context],
            "checklist": list(self.checklist),
            "evidence": [item.to_dict() for item in self.evidence],
            "review_issues": [issue.to_dict() for issue in self.review_issues],
            "group_id": self.group_id,
        }
