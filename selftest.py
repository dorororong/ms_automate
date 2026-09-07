"""파이프라인 자가 점검. Outlook에 아무것도 만들지 않습니다.

    python selftest.py            # 전 구간 점검 (AI 호출 포함)
    python selftest.py --offline  # AI 호출 없이 구조만 점검
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from tempfile import TemporaryDirectory
import time

SAMPLE = """담임 선생님께 부탁드립니다.

1. 내일 종례시간까지 디벗 충전 확인 부탁드립니다.
2. 학급함에 2학기 상담주간 가정통신문을 넣어두었으니 내일 조회 때 학생들에게 배부해 주세요.
3. 8월 27일(목)까지 9월 월중계획을 시트에 입력 부탁드립니다.
4. 8월 25일(화) 15:00 교직원 회의 및 연수가 있습니다. 장소는 시청각실입니다.
5. 매주 화,목요일은 재활용 분리수거날이니 학생 지도 부탁드립니다."""

REFERENCE_DATE = "2026-08-24"


def structured_item(**changes: object) -> dict[str, object]:
    """Build one complete Solar response item for parser regression checks."""

    item: dict[str, object] = {
        "scope": "업무",
        "type": "todo",
        "title": "자료 제출",
        "detail": "자료를 제출합니다.",
        "intent": "submit",
        "applicability": "required",
        "condition": None,
        "source_state": "pending",
        "at": None,
        "due": "2026-09-10",
        "due_time": None,
        "all_day": False,
        "location": None,
        "reminder_minutes": None,
        "repeat_freq": "none",
        "repeat_detail": "",
        "temporal_context": [],
        "checklist": [],
        "evidence": [{"segment_id": "current", "field": "title", "quote": "자료 제출"}],
        "review_issues": [],
        "group_id": "g1",
        "urgency": "보통",
    }
    item.update(changes)
    return item


def check(label: str, ok: bool, detail: str = "") -> bool:
    print(f"  {'OK  ' if ok else 'FAIL'} {label}" + (f" - {detail}" if detail else ""))
    return ok


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--offline", action="store_true", help="AI 호출 없이 점검")
    args = parser.parse_args()
    results: list[bool] = []

    print("\n[1] 모듈 / 설정")
    from models import WorkItem
    from upstage_classifier import REASONING_EFFORT, has_api_key, load_profile
    results.append(check("프로필 로드", bool(load_profile()), load_profile()[:50]))
    results.append(check("추론 비활성", REASONING_EFFORT == "none", f"reasoning_effort={REASONING_EFFORT}"))
    print(f"  INFO API 키 {'있음' if has_api_key() else '없음'}")

    print("\n[2] 2차 분류 파생 규칙")
    for scope, at, due, expected, target in [
        ("업무", None, None, "todo", "todo"),
        ("업무", "2026-09-10T14:00:00", None, "todo_with_date", "calendar"),
        ("업무", None, "2026-09-15", "todo_with_due_date", "todo"),
        ("참조", "2026-09-10T14:00:00", None, "none", None),
    ]:
        item = WorkItem(type="todo", scope=scope, title="t", start=at, due=due)
        results.append(check(f"{scope} at={bool(at)} due={bool(due)}",
                             item.action == expected and item.outlook_target == target,
                             f"{item.action} → {item.outlook_target}"))

    print("\n[3] 요일 파싱")
    from outlook_adapter import parse_weekday_mask
    for text, expected in [("화,목", 20), ("매일 아침", 0), ("매월 1일", 0), ("월요일", 2)]:
        results.append(check(f"{text!r}", parse_weekday_mask(text) == expected))

    print("\n[4] 전역 단축키 등록")
    from hotkey import HotkeyError, HotkeyListener
    listener = HotkeyListener(lambda: None, key_code=0x58, label="Ctrl+Shift+X")
    try:
        listener.start(); listener.stop()
        results.append(check("Ctrl+Shift+X 등록/해제", True))
    except HotkeyError as exc:
        results.append(check("Ctrl+Shift+X 등록/해제", False, str(exc)))

    print("\n[5] 클립보드")
    import selection
    before = selection.read_clipboard_text()
    selection.write_clipboard_text("__selftest__")
    ok = selection.read_clipboard_text() == "__selftest__"
    if before:
        selection.write_clipboard_text(before)
    results.append(check("읽기/쓰기/복원", ok))

    print("\n[6] Outlook 연결 (읽기 전용)")
    from outlook_adapter import OutlookAdapter, OutlookUnavailableError
    try:
        name = OutlookAdapter().test_connection()
        results.append(check("클래식 Outlook COM", True, name))
    except OutlookUnavailableError as exc:
        results.append(check("클래식 Outlook COM", False, str(exc)[:60]))

    print("\n[7] 행동 단위·의미 필드")
    from classifier import classify_text
    toeic_text = """안녕하세요.
토익 시험 감독관을 선착순으로 모집합니다.
1. 일시: 2026. 8. 23(일) 08:30
2. 수당: 8만원
3. 모집 인원: 12명
※ 당일 오전 10시까지 접수 후 마감됩니다."""
    toeic = classify_text(toeic_text)
    toeic_item = toeic.work_items[0] if toeic.work_items else None
    results.append(check("모집 공지 1개 행동", len(toeic.work_items) == 1))
    results.append(check(
        "행사일/마감시각 문맥 보존",
        bool(toeic_item)
        and toeic_item.due is None
        and toeic_item.due_time == "10:00"
        and any(c.role == "event_context" for c in toeic_item.temporal_context),
    ))
    results.append(check(
        "조건부 기본 선택 해제",
        bool(toeic_item)
        and toeic_item.applicability == "conditional"
        and not toeic_item.selected
        and any(i.code == "ambiguous_deadline_date" for i in toeic_item.review_issues),
    ))

    from models import Evidence, TemporalContext
    from storage import Database
    with TemporaryDirectory() as temp_dir:
        db = Database(Path(temp_dir) / "semantic.sqlite3")
        stored = WorkItem(
            type="todo",
            title="의미 보존 테스트",
            due_time="15:00",
            intent="submit",
            applicability="conditional",
            condition="해당하는 경우",
            temporal_context=[TemporalContext(role="event_context", date="2026-09-10")],
            evidence=[Evidence(field="title", quote="의미 보존 테스트")],
        )
        input_id = db.create_input(
            "의미 보존 테스트",
            [stored],
            context={"reference_date": "2026-09-07"},
        )
        loaded = db.get_work_items(input_id)[0]
        results.append(check(
            "DB v2 의미 필드 재로딩",
            loaded.due_time == "15:00"
            and loaded.intent == "submit"
            and loaded.applicability == "conditional"
            and bool(loaded.temporal_context)
            and bool(loaded.evidence),
        ))

    print("\n[8] Solar 응답 하네스·등록 안전장치")
    from upstage_classifier import (
        MAX_API_ATTEMPTS,
        REQUEST_TIMEOUT_SECONDS,
        UpstageAPIError,
        _to_classification,
    )
    results.append(check(
        "API 제한시간·재시도 기준",
        REQUEST_TIMEOUT_SECONDS >= 50 and MAX_API_ATTEMPTS >= 2,
        f"timeout={REQUEST_TIMEOUT_SECONDS}s, attempts={MAX_API_ATTEMPTS}",
    ))
    try:
        _to_classification({}, "자료 제출")
    except UpstageAPIError:
        results.append(check("items 누락 응답 거부", True))
    else:
        results.append(check("items 누락 응답 거부", False))

    malformed = structured_item()
    malformed.pop("scope")
    try:
        _to_classification({"items": [malformed]}, "자료 제출")
    except UpstageAPIError:
        results.append(check("필수 필드 누락 응답 거부", True))
    else:
        results.append(check("필수 필드 누락 응답 거부", False))

    date_only = _to_classification(
        {"items": [structured_item(
            title="행사 참석",
            type="calendar",
            at="2026-09-10",
            due=None,
            intent="attend",
            evidence=[{"segment_id": "current", "field": "at", "quote": "9월 10일"}],
        )]},
        "9월 10일 행사 참석",
    ).work_items[0]
    results.append(check(
        "날짜만 있는 Calendar 차단",
        date_only.start is None
        and any(issue.code == "missing_execution_time" for issue in date_only.review_issues)
        and not date_only.selected,
    ))

    guessed_time = _to_classification(
        {"items": [structured_item(
            title="교시 임장",
            type="calendar",
            at="2026-09-10T09:00:00",
            due=None,
            intent="attend",
            evidence=[{"segment_id": "current", "field": "at", "quote": "2교시"}],
        )]},
        "9월 10일 2교시 임장",
    ).work_items[0]
    results.append(check(
        "원문에 없는 시각 추정 차단",
        guessed_time.start is None
        and any(issue.code == "execution_time_not_in_source" for issue in guessed_time.review_issues)
        and not guessed_time.selected,
    ))

    all_day = _to_classification(
        {"items": [structured_item(
            title="개교기념일",
            type="calendar",
            at="2026-09-10",
            due=None,
            intent="attend",
            all_day=True,
            evidence=[{"segment_id": "current", "field": "at", "quote": "9월 10일"}],
        )]},
        "9월 10일 개교기념일 종일",
    ).work_items[0]
    results.append(check(
        "명시된 종일 Calendar 허용",
        all_day.start == "2026-09-10T00:00:00"
        and all_day.all_day
        and all_day.can_register,
    ))

    conditional = _to_classification(
        {"items": [structured_item(
            title="희망자 신청",
            due=None,
            intent="apply",
            applicability="conditional",
            evidence=[{"segment_id": "current", "field": "title", "quote": "희망자 신청"}],
        )]},
        "희망자 신청",
    ).work_items[0]
    results.append(check(
        "조건 누락 조건부 항목 차단",
        not conditional.selected
        and any(issue.code == "missing_condition" for issue in conditional.review_issues),
    ))

    contradictory = _to_classification(
        {"items": [structured_item(
            title="해당자 제출",
            condition="해당자만 제출",
            evidence=[{"segment_id": "current", "field": "condition", "quote": "해당자만 제출"}],
        )]},
        "해당자만 제출",
    ).work_items[0]
    results.append(check(
        "조건이 있는 필수 항목은 조건부로 보류",
        contradictory.applicability == "conditional"
        and not contradictory.selected,
    ))

    from service import WorkflowService
    fallback_item = WorkItem(type="todo", title="규칙 결과 확인")
    WorkflowService._mark_ai_fallback([fallback_item])
    results.append(check(
        "AI 실패 결과는 명시 선택 필요",
        not fallback_item.selected
        and any(issue.code == "ai_fallback" for issue in fallback_item.review_issues),
    ))

    class FailingOutlook:
        def __init__(self) -> None:
            self.saved = 0

        def read_entries(self, limit: int = 200) -> list[object]:
            raise OutlookUnavailableError("selftest duplicate read failure")

        def save_work_item(self, item: WorkItem) -> str | None:
            self.saved += 1
            return "unexpected"

    with TemporaryDirectory() as temp_dir:
        failing_outlook = FailingOutlook()
        workflow = WorkflowService(
            database=Database(Path(temp_dir) / "guard.sqlite3"),
            outlook=failing_outlook,
        )
        outcome = workflow.record([WorkItem(type="todo", title="중복 확인 필수")])
        results.append(check(
            "중복 조회 실패 시 등록 보류",
            outcome.blocked and outcome.saved == 0 and failing_outlook.saved == 0,
        ))

    if not args.offline:
        print("\n[7] Solar Pro 4 구조화 (Outlook 기록 없음)")
        from upstage_classifier import classify_with_upstage
        try:
            t = time.monotonic()
            result = classify_with_upstage(SAMPLE, reference_date=REFERENCE_DATE)
            elapsed = time.monotonic() - t
            results.append(check("API 응답", True, f"{elapsed:.1f}s, {len(result.work_items)}항목"))
            results.append(check("항목 3건 이상 추출", len(result.work_items) >= 3))
            results.append(check("반복 항목 인식",
                                 any(i.repeat_freq != "none" for i in result.work_items)))
            print()
            for n, i in enumerate(result.work_items, 1):
                extra = f" repeat={i.repeat_freq}/{i.repeat_detail}" if i.repeat_freq != "none" else ""
                print(f"    {n}. [{i.scope}/{i.action}] {i.title}")
                print(f"       at={i.start} due={i.due} loc={i.location}{extra} → {i.outlook_target}")
        except Exception as exc:
            results.append(check("API 응답", False, str(exc)[:70]))

    passed = sum(results)
    print(f"\n{'='*60}\n결과: {passed}/{len(results)} 통과")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    sys.exit(main())
