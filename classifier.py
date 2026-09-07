"""Deterministic Korean text classifier for the first Outlook MVP.

This module intentionally does not call an AI model. It only creates a useful,
testable draft that the user can edit before saving to Outlook.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta

from models import (
    CALENDAR,
    Evidence,
    ReviewIssue,
    TemporalContext,
    TODO,
    WorkItem,
)


CALENDAR_KEYWORDS = (
    "회의",
    "미팅",
    "행사",
    "교육",
    "연수",
    "설명회",
    "모임",
    "수업",
    "체험학습",
    "면담",
    "발표",
    "공연",
    "출장",
    "워크숍",
    "워크샵",
    "일정",
)
TODO_KEYWORDS = (
    "제출",
    "작성",
    "신청",
    "확인",
    "준비",
    "회신",
    "마감",
    "등록",
    "공유",
    "보내",
    "업로드",
    "수합",
    "취합",
    "납부",
    "예약",
    "완료",
    "처리",
    "접수",
    "지원",
    "응답",
)

_ISO_DATE_RE = re.compile(
    r"(?<!\d)(?P<year>\d{4})\s*[./-]\s*(?P<month>\d{1,2})\s*[./-]\s*(?P<day>\d{1,2})(?!\d)"
)
_KOREAN_DATE_RE = re.compile(
    r"(?P<year>\d{4})\s*년\s*(?P<month>\d{1,2})\s*월\s*(?P<day>\d{1,2})\s*일?"
)
_TIME_RE = re.compile(
    r"(?P<ampm>오전|오후|AM|PM)?\s*"
    r"(?P<hour>\d{1,2})"
    r"(?:\s*시(?:\s*(?P<minute>\d{1,2})\s*분?)?|\s*:\s*(?P<colon_minute>\d{2}))",
    re.IGNORECASE,
)
_LOCATION_RE = re.compile(
    r"(?:장소|위치|회의실|location)\s*[:：=]\s*"
    r"(?P<location>.*?)(?=\s+(?:알림|리마인더|reminder)\b|[,，;；|]|$)",
    re.IGNORECASE,
)
_REMINDER_PATTERNS = (
    re.compile(
        r"(?:알림|리마인더|reminder)\s*[:：=]?\s*"
        r"(?P<value>\d+)\s*(?P<unit>분|시간)(?:\s*전)?",
        re.IGNORECASE,
    ),
    re.compile(
        r"(?P<value>\d+)\s*(?P<unit>분|시간)\s*전\s*"
        r"(?:알림|리마인더|reminder)",
        re.IGNORECASE,
    ),
)
_FIELD_LINE = re.compile(
    r"^\s*(?P<key>[^=:\n]+?)\s*(?:=|:)\s*(?P<value>.*)\s*$"
)
_STRUCTURED_ALIASES = {
    "title": "title",
    "제목": "title",
    "target": "target",
    "대상": "target",
    "type": "target",
    "분류": "target",
    "start": "start",
    "시작": "start",
    "시작일": "start",
    "end": "end",
    "종료": "end",
    "종료일": "end",
    "due": "due",
    "마감": "due",
    "마감일": "due",
    "due_time": "due_time",
    "마감시각": "due_time",
    "마감시간": "due_time",
    "memo": "memo",
    "메모": "memo",
    "description": "memo",
    "설명": "memo",
    "location": "location",
    "장소": "location",
    "위치": "location",
    "회의실": "location",
    "category": "category",
    "범주": "category",
    "카테고리": "category",
    "all_day": "all_day",
    "allday": "all_day",
    "종일": "all_day",
    "종일일정": "all_day",
    "duration_minutes": "duration_minutes",
    "duration": "duration_minutes",
    "기간": "duration_minutes",
    "reminder_minutes": "reminder_minutes",
    "알림분": "reminder_minutes",
    "reminder": "reminder",
    "알림": "reminder",
    "work_id": "work_id",
}


@dataclass
class ClassificationResult:
    work_items: list[WorkItem]
    ignored: list[str]


def _parse_date(text: str) -> date | None:
    match = _KOREAN_DATE_RE.search(text) or _ISO_DATE_RE.search(text)
    if not match:
        return None
    try:
        return date(int(match.group("year")), int(match.group("month")), int(match.group("day")))
    except ValueError:
        return None


def _parse_time(match: re.Match[str], default_ampm: str | None = None) -> tuple[int, int] | None:
    ampm = (match.group("ampm") or default_ampm or "").lower()
    hour = int(match.group("hour"))
    minute = int(match.group("minute") or match.group("colon_minute") or 0)
    if minute > 59:
        return None
    if ampm in {"오후", "pm"} and hour < 12:
        hour += 12
    if ampm in {"오전", "am"} and hour == 12:
        hour = 0
    if hour > 23:
        return None
    return hour, minute


def _parse_start_end(text: str) -> tuple[str | None, str | None]:
    event_date = _parse_date(text)
    time_matches = list(_TIME_RE.finditer(text))
    if event_date is None or not time_matches:
        return None, None

    first = _parse_time(time_matches[0])
    if first is None:
        return None, None
    start_dt = datetime.combine(event_date, datetime.min.time()).replace(hour=first[0], minute=first[1])

    end_value: str | None = None
    if len(time_matches) > 1:
        second = _parse_time(time_matches[1], time_matches[0].group("ampm"))
        if second is not None:
            end_dt = datetime.combine(event_date, datetime.min.time()).replace(hour=second[0], minute=second[1])
            if end_dt > start_dt:
                end_value = end_dt.isoformat(timespec="seconds")
    return start_dt.isoformat(timespec="seconds"), end_value


def _parse_location(text: str) -> str | None:
    match = _LOCATION_RE.search(text)
    if not match:
        return None
    location = match.group("location").strip(" \t\r\n.,。")
    return location[:500] or None


def _parse_reminder_minutes(text: str) -> int | None:
    for pattern in _REMINDER_PATTERNS:
        match = pattern.search(text)
        if not match:
            continue
        value = int(match.group("value"))
        if match.group("unit") == "시간":
            value *= 60
        return value if 0 <= value <= 10080 else None
    return None


def _structured_key(value: str) -> str | None:
    normalized = re.sub(r"[\s_-]+", "", value.strip().casefold())
    for alias, canonical in _STRUCTURED_ALIASES.items():
        if re.sub(r"[\s_-]+", "", alias.casefold()) == normalized:
            return canonical
    return None


def _parse_structured_datetime(value: str, default_date: date | None = None) -> datetime | None:
    text = value.strip().replace("Z", "+00:00")
    if not text:
        return None

    time_match = _TIME_RE.search(text)
    parsed_date = _parse_date(text)
    if parsed_date is not None and time_match is not None:
        parsed_time = _parse_time(time_match)
        if parsed_time is not None:
            return datetime.combine(parsed_date, datetime.min.time()).replace(
                hour=parsed_time[0], minute=parsed_time[1]
            )
    if parsed_date is not None and time_match is None:
        return datetime.combine(parsed_date, datetime.min.time())
    if default_date is not None and time_match is not None:
        parsed_time = _parse_time(time_match)
        if parsed_time is not None:
            return datetime.combine(default_date, datetime.min.time()).replace(
                hour=parsed_time[0], minute=parsed_time[1]
            )

    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        return None
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone().replace(tzinfo=None)
    return parsed.replace(second=0, microsecond=0)


def _parse_structured_bool(value: str | None) -> bool:
    if not value:
        return False
    return value.strip().casefold() in {"1", "true", "yes", "y", "on", "예", "네"}


def _parse_structured_int(value: str | None) -> int | None:
    if not value:
        return None
    try:
        parsed = int(value.strip())
    except ValueError:
        return None
    return parsed if 0 <= parsed <= 10080 else None


def _classify_structured_text(text: str) -> ClassificationResult | None:
    """Read the command format used by the reference Outlook project."""

    fields: dict[str, str] = {}
    memo_parts: list[str] = []
    found_field = False
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        match = _FIELD_LINE.match(line)
        if not match:
            if found_field:
                memo_parts.append(line)
            continue
        key = _structured_key(match.group("key"))
        if key is None:
            if found_field:
                memo_parts.append(line)
            continue
        found_field = True
        value = match.group("value").strip()
        if key == "memo":
            if value:
                memo_parts.append(value)
        else:
            fields[key] = value

    title = fields.get("title", "").strip()
    if not title:
        return None

    target = fields.get("target", "").strip().casefold()
    if target not in {"calendar", "todo", "both"}:
        target = ""

    start_raw = fields.get("start", "")
    start_dt = _parse_structured_datetime(start_raw)
    due_date = _parse_date(fields.get("due", "")) if fields.get("due") else None
    due_time = None
    if fields.get("due_time"):
        due_time_match = _TIME_RE.search(fields["due_time"])
        if due_time_match:
            parsed_due_time = _parse_time(due_time_match)
            if parsed_due_time is not None:
                due_time = f"{parsed_due_time[0]:02d}:{parsed_due_time[1]:02d}"
    used_due_as_start = False
    if start_dt is None and target in {"calendar", "both"} and due_date is not None:
        start_dt = datetime.combine(due_date, datetime.min.time())
        used_due_as_start = True

    all_day = _parse_structured_bool(fields.get("all_day"))
    if start_dt is not None and (used_due_as_start or (start_raw and not _TIME_RE.search(start_raw))):
        all_day = True

    end_dt = None
    if fields.get("end"):
        end_dt = _parse_structured_datetime(
            fields["end"], default_date=start_dt.date() if start_dt else None
        )
    duration = _parse_structured_int(fields.get("duration_minutes"))
    if end_dt is None and duration and start_dt is not None and not all_day:
        end_dt = start_dt + timedelta(minutes=duration)

    reminder_minutes = _parse_structured_int(fields.get("reminder_minutes"))
    reminder_value = None
    if fields.get("reminder"):
        reminder_dt = _parse_structured_datetime(
            fields["reminder"], default_date=start_dt.date() if start_dt else None
        )
        if reminder_dt is not None:
            reminder_value = reminder_dt.isoformat(timespec="seconds")
            if reminder_minutes is None and start_dt is not None:
                delta = int((start_dt - reminder_dt).total_seconds() / 60)
                if 0 <= delta <= 10080:
                    reminder_minutes = delta

    location = fields.get("location", "").strip() or None
    category = fields.get("category", "").strip() or None
    description = " ".join(memo_parts).strip() or title
    start_value = start_dt.isoformat(timespec="seconds") if start_dt else None
    end_value = end_dt.isoformat(timespec="seconds") if end_dt else None
    due_value = due_date.isoformat() if due_date else None

    if not target:
        target = CALENDAR if start_dt or any(keyword in title for keyword in CALENDAR_KEYWORDS) else TODO

    def build(item_type: str) -> WorkItem:
        return WorkItem(
            type=item_type,
            title=title[:240],
            start=start_value if item_type == CALENDAR else None,
            end=end_value if item_type == CALENDAR else None,
            due=due_value if item_type == TODO else None,
            due_time=due_time if item_type == TODO else None,
            description=description,
            location=location if item_type == CALENDAR else None,
            category=category,
            all_day=all_day if item_type == CALENDAR else False,
            reminder_minutes=reminder_minutes if item_type == CALENDAR else None,
            reminder=reminder_value,
        )

    if target == "both":
        return ClassificationResult(work_items=[build(CALENDAR), build(TODO)], ignored=[])
    return ClassificationResult(work_items=[build(target)], ignored=[])


def _clean_title(line: str) -> str:
    cleaned = re.sub(r"^\s*(?:[-*•]|\d+[.)])\s*", "", line)
    cleaned = re.sub(
        r"^\s*(?:calendar|todo|ignore|일정|할\s*일|업무)\s*[:：]\s*",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )
    cleaned = _LOCATION_RE.sub(" ", cleaned)
    for pattern in _REMINDER_PATTERNS:
        cleaned = pattern.sub(" ", cleaned)
    cleaned = _KOREAN_DATE_RE.sub(" ", cleaned)
    cleaned = _ISO_DATE_RE.sub(" ", cleaned)
    cleaned = _TIME_RE.sub(" ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned.strip(" \t\r\n,，;；:：-|/")
    return (cleaned or line.strip())[:240]


def _split_candidates(text: str) -> list[str]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if len(lines) == 1:
        # Do not split Korean/ISO dates such as `2026. 8. 23` at the
        # punctuation between numeric components.
        lines = [
            part.strip()
            for part in re.split(r"(?<=[!?。])\s+(?!\d)", lines[0])
            if part.strip()
        ]
    return lines


def _classify_recruitment_notice(text: str) -> ClassificationResult | None:
    """Conservative fallback for 모집/선착순 notices.

    A recruitment notice contains several dates and numbers, but its user
    action is normally one conditional application/reply. Keep related event
    and relative-deadline information on that one candidate instead of making
    one Outlook item per date.
    """

    if not re.search(r"모집|선착순|희망자|지원자", text):
        return None
    action_line = next(
        (
            line.strip()
            for line in text.splitlines()
            if re.search(r"모집|선착순|신청|지원", line)
        ),
        "",
    )
    if not action_line:
        return None
    subject = re.sub(
        r"(?:을|를)?\s*(?:선착순으로\s*)?(?:모집|신청|지원)(?:합니다|받습니다|중입니다)?[.!。]?$",
        "",
        action_line,
    )
    subject = re.sub(r"\s+", " ", subject).strip(" .,:：-—")
    subject = re.sub(r"(?:은|는|이|가)$", "", subject).strip()
    title = f"{subject} 신청" if subject else "모집 신청"
    event_date = _parse_date(text)
    first_time = _TIME_RE.search(text)
    event_time: str | None = None
    if first_time:
        parsed_time = _parse_time(first_time)
        if parsed_time is not None:
            event_time = f"{parsed_time[0]:02d}:{parsed_time[1]:02d}"

    contexts: list[TemporalContext] = []
    if event_date or event_time:
        contexts.append(
            TemporalContext(
                role="event_context",
                label="모집 대상 행사 일시",
                raw_text=(
                    next(
                        (line.strip() for line in text.splitlines() if "일시" in line),
                        "",
                    )
                    or (event_date.isoformat() if event_date else "")
                ),
                date=event_date.isoformat() if event_date else None,
                time=event_time,
                precision="datetime" if event_date and event_time else "date" if event_date else "unknown",
                resolution="explicit" if event_date or event_time else "unresolved",
            )
        )

    relative_deadline = re.search(r"(?:당일|그날)[^\n]{0,40}", text)
    deadline_time: str | None = None
    if relative_deadline:
        deadline_match = _TIME_RE.search(relative_deadline.group(0))
        if deadline_match:
            parsed_time = _parse_time(deadline_match)
            if parsed_time is not None:
                deadline_time = f"{parsed_time[0]:02d}:{parsed_time[1]:02d}"
        contexts.append(
            TemporalContext(
                role="deadline",
                label="상대적 접수 마감",
                raw_text=relative_deadline.group(0).strip(),
                date=None,
                time=deadline_time,
                precision="unknown",
                resolution="unresolved",
            )
        )

    detail_lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip() and not line.strip().startswith("안녕하세요")
    ]
    detail = "\n".join(detail_lines)[:4000]
    evidence = [Evidence(field="intent", quote=action_line[:500])]
    if relative_deadline:
        evidence.append(Evidence(field="deadline", quote=relative_deadline.group(0).strip()[:500]))
    issues: list[ReviewIssue] = []
    if relative_deadline:
        issues.append(
            ReviewIssue(
                code="ambiguous_deadline_date",
                field="due",
                message="‘당일/그날’의 기준 날짜가 원문에서 확정되지 않았습니다.",
                blocking=True,
            )
        )
    item = WorkItem(
        type=TODO,
        scope="업무",
        title=title[:240],
        description=detail or title,
        due_time=deadline_time,
        intent="apply",
        applicability="conditional",
        condition="모집 대상에 해당하고 참여를 희망하는 경우",
        source_state="pending",
        temporal_context=contexts,
        evidence=evidence,
        review_issues=issues,
        group_id="group-1",
        selected=False,
    )
    return ClassificationResult(work_items=[item], ignored=[])


def classify_text(text: str) -> ClassificationResult:
    """Create calendar/todo drafts and retain unclassified lines as ignore."""

    structured_result = _classify_structured_text(text)
    if structured_result is not None:
        return structured_result

    recruitment_result = _classify_recruitment_notice(text)
    if recruitment_result is not None:
        return recruitment_result

    work_items: list[WorkItem] = []
    ignored: list[str] = []
    for line in _split_candidates(text):
        if re.search(r"(?:완료|처리했|보냈|전달했|제출했|확인했)", line) and not re.search(
            r"(?:해야|부탁|필요|재제출|수정|누락)", line
        ):
            # A completion/update sentence is source context, not a fresh
            # pending task. Keep it in the collapsed ignored area.
            ignored.append(line)
            continue
        title = _clean_title(line)
        has_calendar_keyword = any(keyword in line for keyword in CALENDAR_KEYWORDS)
        has_todo_keyword = any(keyword in line for keyword in TODO_KEYWORDS)
        start, end = _parse_start_end(line)
        due_date = _parse_date(line)
        due = due_date.isoformat() if due_date and has_todo_keyword else None

        if start or has_calendar_keyword:
            item_type = CALENDAR
        elif has_todo_keyword:
            item_type = TODO
        else:
            ignored.append(line)
            continue

        # A date-only event is kept as an unresolved Calendar candidate. Do
        # not invent midnight or silently turn an unknown time into all-day.
        all_day = False
        temporal_context: list[TemporalContext] = []
        review_issues: list[ReviewIssue] = []
        if item_type == CALENDAR and start is None and due_date is not None:
            temporal_context.append(
                TemporalContext(
                    role="execution",
                    label="날짜만 있는 수행 시점",
                    raw_text=line,
                    date=due_date.isoformat(),
                    precision="date",
                    resolution="explicit",
                )
            )
            review_issues.append(
                ReviewIssue(
                    code="missing_execution_time",
                    field="start",
                    message="날짜는 있지만 시각이 없어 종일 여부를 확인해야 합니다.",
                    blocking=True,
                )
            )

        # A date without an explicit year is intentionally not interpreted.
        # This keeps uncertain dates as null rather than inventing a year.
        work_items.append(
            WorkItem(
                type=item_type,
                title=title,
                start=start,
                end=end,
                due=due,
                description=line,
                location=_parse_location(line),
                all_day=all_day,
                reminder_minutes=_parse_reminder_minutes(line),
                intent="attend" if item_type == CALENDAR else "other",
                temporal_context=temporal_context,
                evidence=[Evidence(field="source", quote=line[:500])],
                review_issues=review_issues,
            )
        )
    return ClassificationResult(work_items=work_items, ignored=ignored)
