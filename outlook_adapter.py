"""Classic Outlook COM adapter using pywin32."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Any

from models import CALENDAR, TODO, WorkItem

try:
    import win32com.client as win32_client
except ImportError:  # pragma: no cover - depends on the host installation
    win32_client = None


class OutlookUnavailableError(RuntimeError):
    """Raised when classic Outlook or pywin32 is unavailable."""


class OutlookOperationError(RuntimeError):
    """Raised when Outlook rejects an item operation."""


@dataclass(frozen=True)
class OutlookEntry:
    """Read-only snapshot of an item in the current Outlook profile."""

    item_type: str
    entry_id: str
    subject: str
    start: datetime | None = None
    end: datetime | None = None
    due: datetime | None = None
    all_day: bool = False
    location: str = ""
    category: str = ""
    reminder: datetime | None = None
    reminder_minutes: int | None = None
    complete: bool | None = None
    body: str = ""


def _supports_property(item: Any, name: str) -> bool:
    try:
        getattr(item, name)
    except (AttributeError, TypeError, RuntimeError):
        return False
    return True


def _read_local_datetime(value: Any) -> datetime | None:
    """Read an Outlook date as the displayed local clock value.

    pywin32 may attach a timezone object whose conversion treats a local
    Outlook value as UTC. Removing that COM timezone metadata preserves the
    clock value the user sees in classic Outlook.
    """

    if value is None:
        return None
    if isinstance(value, datetime):
        result = datetime(
            value.year,
            value.month,
            value.day,
            value.hour,
            value.minute,
        )
        # Outlook uses 4501-01-01 as a sentinel for a Task with no due date.
        if result.year < 1900 or result.year > 2100:
            return None
        return result
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    return None


def _read_datetime_property(item: Any, *names: str) -> datetime | None:
    for name in names:
        try:
            value = getattr(item, name)
        except Exception:
            continue
        result = _read_local_datetime(value)
        if result is not None:
            return result
    return None


def _read_text_property(item: Any, name: str) -> str:
    try:
        value = getattr(item, name)
    except Exception:
        return ""
    return str(value or "").strip()


def _com_local_datetime(value: datetime) -> datetime:
    """Compensate for pywin32's local-time conversion for TaskItem dates."""

    if value.tzinfo is not None:
        value = value.astimezone().replace(tzinfo=None)
    local_value = value.astimezone()
    offset = local_value.utcoffset() or timedelta(0)
    return value + offset


def _set_appointment_start(item: Any, start: datetime) -> None:
    # Outlook exposes StartUTC/EndUTC for appointments. With pywin32, passing
    # the user's local naive datetime through these properties preserves the
    # displayed local time; older/fake objects use Start/End instead.
    if _supports_property(item, "StartUTC"):
        item.StartUTC = start
    else:
        item.Start = start


def _set_appointment_times(item: Any, start: datetime, end: datetime) -> None:
    if _supports_property(item, "StartUTC") and _supports_property(item, "EndUTC"):
        item.StartUTC = start
        item.EndUTC = end
    else:
        item.Start = start
        item.End = end


def parse_outlook_datetime(value: str | date | datetime) -> datetime:
    if isinstance(value, datetime):
        result = value
    elif isinstance(value, date):
        result = datetime.combine(value, time.min)
    else:
        raw = str(value).strip().replace("Z", "+00:00")
        try:
            result = datetime.fromisoformat(raw)
        except ValueError:
            for fmt in (
                "%Y-%m-%d %H:%M",
                "%Y/%m/%d %H:%M",
                "%Y.%m.%d %H:%M",
                "%Y-%m-%d",
                "%Y/%m/%d",
                "%Y.%m.%d",
            ):
                try:
                    result = datetime.strptime(raw, fmt)
                    break
                except ValueError:
                    continue
            else:
                raise ValueError(f"날짜/시간 형식을 해석할 수 없습니다: {value}")
    if result.tzinfo is not None:
        result = result.astimezone().replace(tzinfo=None)
    return result.replace(second=0, microsecond=0)


def parse_outlook_due(value: str | date | datetime) -> datetime:
    if isinstance(value, datetime):
        return value.replace(hour=0, minute=0, second=0, microsecond=0)
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    raw = str(value).strip().replace("/", "-").replace(".", "-")
    return datetime.strptime(raw, "%Y-%m-%d")


def _outlook_body(work_item: WorkItem) -> str:
    """Keep semantic context visible in Outlook without changing its dates."""

    lines: list[str] = []
    if work_item.description.strip():
        lines.append(work_item.description.strip())
    if work_item.condition:
        lines.append(f"조건: {work_item.condition}")
    if work_item.due_time:
        if work_item.due:
            lines.append(f"마감 시각: {work_item.due_time}")
        else:
            lines.append(f"마감 시각(날짜 확인 필요): {work_item.due_time}")
    if work_item.checklist:
        lines.append("체크리스트:")
        lines.extend(f"- {value}" for value in work_item.checklist)
    for context in work_item.temporal_context:
        label = context.label or context.role
        value = context.raw_text
        if context.date or context.time:
            value = " ".join(part for part in (context.date, context.time) if part)
        if value:
            lines.append(f"관련 {label}: {value}")
    return "\n".join(lines)


# olRecurrenceType
RECUR_DAILY = 0
RECUR_WEEKLY = 1
RECUR_MONTHLY = 2
# olDaysOfWeek 비트마스크
WEEKDAY_BITS = {
    "일": 1, "월": 2, "화": 4, "수": 8, "목": 16, "금": 32, "토": 64,
}
_FREQ_TO_TYPE = {"daily": RECUR_DAILY, "weekly": RECUR_WEEKLY, "monthly": RECUR_MONTHLY}


# 낱글자 요일만 잡는다. "매일"·"평일"·"매월 1일"의 '일'을 일요일로 오인하면
# 엉뚱한 반복 일정이 만들어지므로 앞이 한글/숫자면 제외한다.
_WEEKDAY_RE = re.compile(r"(?<![가-힣0-9])([월화수목금토일])(?:요일)?(?![가-힣])")


def parse_weekday_mask(detail: str) -> int:
    """'화,목' -> olDaysOfWeek 비트마스크. 인식 못하면 0."""

    mask = 0
    for name in _WEEKDAY_RE.findall(detail or ""):
        mask |= WEEKDAY_BITS[name]
    return mask


def _apply_recurrence(item: Any, work_item: WorkItem, start: datetime) -> None:
    """반복 규칙을 적용한다.

    주의: TaskItem 에서 GetRecurrencePattern() 을 호출하는 순간 IsRecurring 이
    True 가 된다. 그래서 반복이 실제로 필요할 때만 호출해야 한다.
    """

    recurrence_type = _FREQ_TO_TYPE.get(work_item.repeat_freq)
    if recurrence_type is None:
        return
    pattern = item.GetRecurrencePattern()
    pattern.RecurrenceType = recurrence_type
    pattern.Interval = 1
    if recurrence_type == RECUR_WEEKLY:
        mask = parse_weekday_mask(work_item.repeat_detail)
        if mask:
            pattern.DayOfWeekMask = mask
    elif recurrence_type == RECUR_MONTHLY:
        pattern.DayOfMonth = start.day
    # PatternStartDate 는 pywin32 가 로컬→UTC 로 변환해 하루 앞당기는 경우가 있다.
    # (KST 09:00 이전 시각이면 전날로 넘어가 첫 인스턴스가 하루 빨라졌다.)
    # TaskItem 날짜와 같은 보정을 적용하고 시각은 날짜 단위로 잘라 넘긴다.
    pattern.PatternStartDate = _com_local_datetime(
        datetime.combine(start.date(), time.min)
    )
    pattern.NoEndDate = True


class OutlookAdapter:
    CALENDAR_ITEM = 1
    TASK_ITEM = 3
    CALENDAR_FOLDER = 9
    TASK_FOLDER = 13
    APPOINTMENT_CLASS = 26
    TASK_CLASS = 48

    def __init__(self) -> None:
        self._outlook: Any | None = None

    def _application(self) -> Any:
        if win32_client is None:
            raise OutlookUnavailableError(
                "pywin32가 설치되지 않았습니다. requirements.txt의 pywin32를 설치하세요."
            )
        if self._outlook is not None:
            return self._outlook
        try:
            self._outlook = win32_client.Dispatch("Outlook.Application")
        except Exception as exc:  # COM errors vary by Outlook installation.
            raise OutlookUnavailableError(
                "클래식 Outlook COM 연결에 실패했습니다. 클래식 Outlook이 설치되어 있는지 확인하세요."
            ) from exc
        return self._outlook

    def test_connection(self) -> str:
        outlook = self._application()
        try:
            session = outlook.GetNamespace("MAPI")
            current_user = getattr(session, "CurrentUser", None)
            user_name = getattr(current_user, "Name", None)
            return str(user_name or "클래식 Outlook 연결됨")
        except Exception:
            return "클래식 Outlook 연결됨"

    def read_entries(self, limit: int = 200) -> list[OutlookEntry]:
        """Read Calendar and Task items from the current Outlook profile.

        This is a read-only operation. It uses the same local COM profile as
        saving items and does not call Microsoft Graph or any external API.
        """

        if limit <= 0:
            raise ValueError("Outlook 내역 조회 limit은 1 이상이어야 합니다.")

        try:
            session = self._application().GetNamespace("MAPI")
            entries: list[OutlookEntry] = []
            folder_specs = (
                (self.CALENDAR_FOLDER, self.APPOINTMENT_CLASS, CALENDAR),
                (self.TASK_FOLDER, self.TASK_CLASS, TODO),
            )
            for folder_id, expected_class, item_type in folder_specs:
                folder = session.GetDefaultFolder(folder_id)
                items = folder.Items
                count = int(items.Count)
                for index in range(1, count + 1):
                    try:
                        item = items.Item(index)
                        if int(getattr(item, "Class", 0)) != expected_class:
                            continue
                        entry_id = _read_text_property(item, "EntryID")
                        subject = _read_text_property(item, "Subject")
                        category = _read_text_property(item, "Categories")
                        body = _read_text_property(item, "Body")
                        if item_type == CALENDAR:
                            reminder_minutes: int | None = None
                            try:
                                if bool(getattr(item, "ReminderSet", False)):
                                    reminder_minutes = int(
                                        getattr(item, "ReminderMinutesBeforeStart")
                                    )
                            except (AttributeError, TypeError, ValueError, RuntimeError):
                                reminder_minutes = None
                            entries.append(
                                OutlookEntry(
                                    item_type=CALENDAR,
                                    entry_id=entry_id,
                                    subject=subject,
                                    start=_read_datetime_property(
                                        item, "StartInStartTimeZone", "Start"
                                    ),
                                    end=_read_datetime_property(
                                        item, "EndInEndTimeZone", "End"
                                    ),
                                    all_day=bool(getattr(item, "AllDayEvent", False)),
                                    location=_read_text_property(item, "Location"),
                                    category=category,
                                    reminder_minutes=reminder_minutes,
                                    body=body,
                                )
                            )
                        else:
                            reminder: datetime | None = None
                            try:
                                if bool(getattr(item, "ReminderSet", False)):
                                    reminder = _read_datetime_property(item, "ReminderTime")
                            except (AttributeError, TypeError, RuntimeError):
                                reminder = None
                            entries.append(
                                OutlookEntry(
                                    item_type=TODO,
                                    entry_id=entry_id,
                                    subject=subject,
                                    due=_read_datetime_property(item, "DueDate"),
                                    category=category,
                                    reminder=reminder,
                                    complete=bool(getattr(item, "Complete", False)),
                                    body=body,
                                )
                            )
                    except Exception:
                        # A folder can contain an item that disappears or is
                        # not fully readable while Outlook is synchronising.
                        continue

            entries.sort(
                key=lambda entry: (
                    entry.start or entry.due or datetime.max,
                    entry.item_type,
                    entry.subject.casefold(),
                )
            )
            return entries[:limit]
        except OutlookOperationError:
            raise
        except Exception as exc:
            raise OutlookOperationError(f"Outlook 내역 읽기 실패: {exc}") from exc

    def _get_item_by_entry_id(self, entry_id: str, item_type: str) -> Any:
        if not entry_id:
            raise ValueError("Outlook EntryID가 비어 있습니다.")
        expected_class = self.APPOINTMENT_CLASS if item_type == CALENDAR else self.TASK_CLASS
        try:
            item = self._application().GetNamespace("MAPI").GetItemFromID(entry_id)
            if int(getattr(item, "Class", 0)) != expected_class:
                raise OutlookOperationError("EntryID의 Outlook 항목 유형이 요청과 다릅니다.")
            return item
        except OutlookOperationError:
            raise
        except Exception as exc:
            raise OutlookOperationError(f"Outlook 항목을 찾을 수 없습니다: {exc}") from exc

    def delete_entry(self, entry_id: str, item_type: str) -> None:
        """Delete one existing Calendar or Task item by EntryID."""

        if item_type not in (CALENDAR, TODO):
            raise ValueError(f"지원하지 않는 Outlook 항목 유형입니다: {item_type}")
        item = self._get_item_by_entry_id(entry_id, item_type)
        try:
            item.Delete()
        except Exception as exc:
            raise OutlookOperationError(f"Outlook 항목 삭제 실패: {exc}") from exc

    def update_entry(self, entry: OutlookEntry, changes: dict[str, Any]) -> None:
        """Update an existing Outlook item selected by the command pipeline."""

        item = self._get_item_by_entry_id(entry.entry_id, entry.item_type)
        try:
            if "new_title" in changes:
                item.Subject = str(changes["new_title"])
            if "new_body" in changes:
                item.Body = str(changes["new_body"])
            if "new_category" in changes:
                item.Categories = str(changes["new_category"])

            if entry.item_type == CALENDAR:
                if "new_location" in changes:
                    item.Location = str(changes["new_location"])
                if "new_date" in changes:
                    if entry.start is None:
                        raise OutlookOperationError(
                            "시작 시간이 없는 Calendar 항목은 날짜를 변경할 수 없습니다."
                        )
                    target_date = date.fromisoformat(str(changes["new_date"])[:10])
                    delta = target_date - entry.start.date()
                    if entry.all_day:
                        start = datetime.combine(target_date, time.min)
                        duration = (
                            entry.end.date() - entry.start.date()
                            if entry.end is not None
                            else timedelta(days=1)
                        )
                        end = start + duration
                    else:
                        start = entry.start + delta
                        end = entry.end + delta if entry.end is not None else None
                    if end is not None:
                        _set_appointment_times(item, start, end)
                    else:
                        _set_appointment_start(item, start)
                if "new_reminder_minutes" in changes:
                    minutes = int(changes["new_reminder_minutes"])
                    if not 0 <= minutes <= 10080:
                        raise ValueError("알림 시간은 0~10080분 사이여야 합니다.")
                    item.ReminderMinutesBeforeStart = minutes
                    item.ReminderSet = True
            else:
                if "new_date" in changes:
                    target_date = date.fromisoformat(str(changes["new_date"])[:10])
                    item.DueDate = _com_local_datetime(
                        datetime.combine(target_date, time.min)
                    )
                if "new_complete" in changes:
                    item.Complete = bool(changes["new_complete"])
                if "new_reminder_minutes" in changes:
                    minutes = int(changes["new_reminder_minutes"])
                    if not 0 <= minutes <= 10080:
                        raise ValueError("알림 시간은 0~10080분 사이여야 합니다.")
                    due = entry.due or datetime.now().replace(microsecond=0)
                    item.ReminderTime = _com_local_datetime(
                        due - timedelta(minutes=minutes)
                    )
                    item.ReminderSet = True
            item.Save()
        except OutlookOperationError:
            raise
        except Exception as exc:
            raise OutlookOperationError(f"Outlook 항목 수정 실패: {exc}") from exc

    def save_work_item(self, work_item: WorkItem) -> str | None:
        work_item.validate()
        target = work_item.outlook_target
        if target is None:
            raise OutlookOperationError(
                "참조 항목은 Outlook에 등록하지 않습니다. 로컬 기록에만 남습니다."
            )
        if target == CALENDAR:
            return self._save_calendar(work_item)
        return self._save_task(work_item)

    def _save_calendar(self, work_item: WorkItem) -> str | None:
        if not work_item.start:
            raise OutlookOperationError(
                "Calendar 항목의 시작 날짜/시간이 없습니다. 날짜와 시간을 확인한 뒤 저장하세요."
            )
        try:
            start = parse_outlook_datetime(work_item.start)
            outlook_item = self._application().CreateItem(self.CALENDAR_ITEM)
            outlook_item.Subject = work_item.title

            if work_item.all_day:
                if work_item.end:
                    end = parse_outlook_datetime(work_item.end)
                else:
                    end = datetime.combine(start.date() + timedelta(days=1), time.min)
                if end <= start:
                    raise OutlookOperationError("종일 일정의 종료일이 시작일보다 빠릅니다.")
                _set_appointment_times(outlook_item, start, end)
                outlook_item.AllDayEvent = True
            else:
                outlook_item.AllDayEvent = False
                if work_item.end:
                    end = parse_outlook_datetime(work_item.end)
                    if end <= start:
                        raise OutlookOperationError("Calendar 종료 시간이 시작 시간보다 빠릅니다.")
                    _set_appointment_times(outlook_item, start, end)
                else:
                    _set_appointment_start(outlook_item, start)
                    outlook_item.Duration = 60

            outlook_item.Body = _outlook_body(work_item)
            if work_item.location:
                outlook_item.Location = work_item.location
            if work_item.category:
                outlook_item.Categories = work_item.category
            reminder_minutes = work_item.reminder_minutes
            if reminder_minutes is None and work_item.reminder:
                reminder_at = parse_outlook_datetime(work_item.reminder)
                reminder_minutes = int((start - reminder_at).total_seconds() / 60)
                if reminder_minutes < 0:
                    raise OutlookOperationError("알림 시각은 일정 시작 시각보다 빨라야 합니다.")
            if reminder_minutes is not None:
                outlook_item.ReminderMinutesBeforeStart = reminder_minutes
                outlook_item.ReminderSet = True
            else:
                outlook_item.ReminderSet = False
            _apply_recurrence(outlook_item, work_item, start)
            outlook_item.Save()
            return str(getattr(outlook_item, "EntryID", "") or "") or None
        except OutlookOperationError:
            raise
        except Exception as exc:
            raise OutlookOperationError(f"Calendar 저장 실패: {exc}") from exc

    def _save_task(self, work_item: WorkItem) -> str | None:
        try:
            outlook_item = self._application().CreateItem(self.TASK_ITEM)
            outlook_item.Subject = work_item.title
            due_start: datetime | None = None
            if work_item.due:
                due_start = parse_outlook_due(work_item.due)
                outlook_item.DueDate = _com_local_datetime(due_start)
            outlook_item.Body = _outlook_body(work_item)
            if work_item.category:
                outlook_item.Categories = work_item.category
            if work_item.reminder:
                outlook_item.ReminderTime = _com_local_datetime(
                    parse_outlook_datetime(work_item.reminder)
                )
                outlook_item.ReminderSet = True
            else:
                outlook_item.ReminderSet = False
            if work_item.repeat_freq != "none":
                _apply_recurrence(
                    outlook_item, work_item, due_start or datetime.now().replace(microsecond=0)
                )
            outlook_item.Save()
            return str(getattr(outlook_item, "EntryID", "") or "") or None
        except Exception as exc:
            raise OutlookOperationError(f"Task 저장 실패: {exc}") from exc
