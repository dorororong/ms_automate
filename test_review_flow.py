"""검토창 회귀 검사: 카드 단위 처리와 메시지 누적.

Tk 위젯을 실제로 만들어 버튼을 눌러 본다. Outlook과 API는 호출하지 않는다.

    python -m unittest test_review_flow -v
"""

from __future__ import annotations

import time
import tkinter as tk
import unittest

from models import WorkItem
from review import ReviewWindow
from service import Analysis, RecordOutcome


class StubService:
    """중복 없음 · 요청한 항목을 그대로 저장하는 최소 서비스."""

    def __init__(self) -> None:
        self.records: list[list[WorkItem]] = []
        self.counter = 0

    def find_conflicts(self, drafts: list[WorkItem]) -> dict[int, str]:
        return {}

    def record(self, drafts: list[WorkItem]) -> RecordOutcome:
        self.records.append(list(drafts))
        outcome = RecordOutcome()
        for item in drafts:
            if not item.selected or item.saved:
                continue
            self.counter += 1
            item.saved = True
            item.outlook_entry_id = f"E{self.counter}"
            outcome.saved += 1
        return outcome


class DuplicateService(StubService):
    def record(self, drafts: list[WorkItem]) -> RecordOutcome:
        outcome = RecordOutcome()
        outcome.duplicates[0] = "기존 Outlook 작업과 같아 제외: 명단 제출"
        return outcome


class FailingService(StubService):
    def record(self, drafts: list[WorkItem]) -> RecordOutcome:
        outcome = RecordOutcome()
        outcome.failed = 1
        outcome.errors[0] = "Outlook 작업 생성 실패"
        return outcome


def make_analysis(input_id: int, titles: list[str]) -> Analysis:
    items = [
        WorkItem(
            id=input_id * 100 + n,
            input_id=input_id,
            type="todo",
            title=title,
            due=f"2026-09-2{n % 10}",
        )
        for n, title in enumerate(titles, start=1)
    ]
    return Analysis(
        input_id=input_id,
        text=f"메시지 {input_id} 본문입니다. " + " / ".join(titles),
        work_items=items,
        source="offline",
        model="rule-based",
    )


class ReviewFlowTest(unittest.TestCase):
    root: tk.Tk

    @classmethod
    def setUpClass(cls) -> None:
        try:
            cls.root = tk.Tk()
        except tk.TclError as exc:  # 화면이 없는 환경
            raise unittest.SkipTest(f"Tk를 열 수 없습니다: {exc}") from exc
        cls.root.withdraw()

    @classmethod
    def tearDownClass(cls) -> None:
        cls.root.destroy()

    def setUp(self) -> None:
        self.finished: list[str] = []

    def pump(self, seconds: float = 0.4) -> None:
        """Tk 가 배치를 마칠 때까지 이벤트를 돌린다."""

        end = time.monotonic() + seconds
        while time.monotonic() < end:
            self.root.update()
            time.sleep(0.01)

    def open_window(self, service: StubService) -> ReviewWindow:
        window = ReviewWindow(self.root, service, on_finished=self.finished.append)
        self.addCleanup(lambda: window.winfo_exists() and window.destroy())
        return window

    def test_cards_register_one_at_a_time(self) -> None:
        service = StubService()
        window = self.open_window(service)
        window.present_analysis(make_analysis(1, ["명단 제출", "회의 참석 준비"]))
        window.set_queue_status(1)

        card = window.cards[0]
        card.register_button.invoke()

        self.assertTrue(card.item.saved)
        self.assertEqual(card.item.outlook_entry_id, "E1")
        self.assertEqual(len(service.records[-1]), 1, "카드 한 건만 보내야 한다")
        self.assertEqual(str(card.register_button.cget("state")), "disabled")
        self.assertFalse(window.cards[1].item.saved, "다른 카드는 건드리지 않는다")

        # 한 메시지에 여러 건이 있어도 나머지를 이어서 등록할 수 있어야 한다.
        second = window.cards[1]
        self.assertTrue(second.is_pending())
        self.assertEqual(str(second.register_button.cget("state")), "normal")
        second.register_button.invoke()
        self.assertTrue(second.item.saved)
        self.assertEqual(second.item.outlook_entry_id, "E2")

    def test_every_item_of_one_message_has_its_own_buttons(self) -> None:
        window = self.open_window(StubService())
        window.present_analysis(make_analysis(1, ["명단 제출", "회의 준비", "연수 신청"]))
        window.set_queue_status(1)

        self.assertEqual(len(window.cards), 3)
        for card in window.cards:
            self.assertEqual(str(card.register_button.cget("state")), "normal")
            self.assertEqual(str(card.dismiss_button.cget("text")), "미등록")

        window.cards[1].dismiss_button.invoke()
        window.cards[0].register_button.invoke()
        window.cards[2].register_button.invoke()

        self.assertTrue(window.cards[0].item.saved)
        self.assertFalse(window.cards[1].item.saved)
        self.assertTrue(window.cards[2].item.saved)

    def test_long_title_is_fully_visible(self) -> None:
        """요약 제목이 길면 줄을 늘려 끝까지 보여야 한다."""

        window = self.open_window(StubService())
        long_title = "2학기 학부모 상담주간 운영 계획 수립 및 상담 희망 조사 결과 취합 제출"
        window.present_analysis(make_analysis(1, [long_title]))
        window.set_queue_status(1)
        self.pump()

        card = window.cards[0]
        self.assertEqual(card.title_text.get("1.0", "end-1c"), long_title)
        self.assertGreater(
            int(card.title_text.cget("height")), 1, "긴 제목은 두 줄 이상이어야 한다"
        )
        self.assertIsNotNone(
            card.title_text.bbox("end-2c"), "제목 마지막 글자가 보여야 한다"
        )

    def test_short_title_stays_one_line(self) -> None:
        window = self.open_window(StubService())
        window.present_analysis(make_analysis(1, ["명단 제출"]))
        window.set_queue_status(1)
        self.pump()

        card = window.cards[0]
        self.assertEqual(int(card.title_text.cget("height")), 1)
        self.assertIsNotNone(card.title_text.bbox("end-2c"))

    def test_editing_the_title_updates_the_draft(self) -> None:
        window = self.open_window(StubService())
        window.present_analysis(make_analysis(1, ["명단 제출"]))
        window.set_queue_status(1)
        card = window.cards[0]

        card.title_text.delete("1.0", "end")
        card.title_text.insert("1.0", "학생 명단 최종 제출")
        card.title_text.event_generate("<KeyRelease>")
        self.pump(0.2)

        self.assertEqual(card.to_work_item().title, "학생 명단 최종 제출")

    def test_dismiss_is_reversible(self) -> None:
        window = self.open_window(StubService())
        window.present_analysis(make_analysis(1, ["명단 제출"]))
        window.set_queue_status(1)
        card = window.cards[0]

        card.dismiss_button.invoke()
        self.assertTrue(card.dismissed)
        self.assertFalse(card.is_pending())
        self.assertEqual(str(card.dismiss_button.cget("text")), "되돌리기")

        card.dismiss_button.invoke()
        self.assertFalse(card.dismissed)
        self.assertTrue(card.is_pending())

    def test_second_message_is_appended_not_queued(self) -> None:
        window = self.open_window(StubService())
        window.present_analysis(make_analysis(1, ["명단 제출"]))
        window.present_analysis(make_analysis(2, ["연수 신청"]))

        self.assertEqual(len(window.groups), 2)
        self.assertEqual(len(window.cards), 2)
        self.assertIs(window.focused, window.groups[0], "초점은 처리 중인 메시지에 남는다")
        self.assertIn("메시지 2", window.groups[1].header.cget("text"))

    def test_focus_follows_the_unfinished_message(self) -> None:
        window = self.open_window(StubService())
        window.present_analysis(make_analysis(1, ["명단 제출"]))
        window.present_analysis(make_analysis(2, ["연수 신청"]))
        window.set_queue_status(1)

        window.cards[0].register_button.invoke()

        self.assertIs(window.focused, window.groups[1])
        self.assertIn("메시지 2 본문", window.input_text.get("1.0", "end-1c"))

    def test_window_closes_only_when_nothing_is_left(self) -> None:
        window = self.open_window(StubService())
        window.present_analysis(make_analysis(1, ["명단 제출"]))
        window.set_queue_status(1)

        window.cards[0].register_button.invoke()
        self.assertTrue(window.winfo_exists(), "분석이 남아 있으면 창을 유지한다")
        self.assertEqual(self.finished, [])

        window.set_queue_status(0)
        self.assertEqual(self.finished, ["registered"])

    def test_duplicate_found_on_click_does_not_block_the_window(self) -> None:
        window = self.open_window(DuplicateService())
        window.present_analysis(make_analysis(1, ["명단 제출"]))
        window.set_queue_status(1)
        card = window.cards[0]

        card.register_button.invoke()

        self.assertFalse(card.item.saved)
        self.assertTrue(card.item.excluded_reason)
        self.assertFalse(card.is_pending(), "제외된 카드는 결정이 끝난 것으로 본다")

        window.set_queue_status(0)
        self.assertEqual(self.finished, ["skipped"])

    def test_failed_registration_can_be_retried(self) -> None:
        window = self.open_window(FailingService())
        window.present_analysis(make_analysis(1, ["보고서 제출"]))
        window.set_queue_status(1)
        card = window.cards[0]

        card.register_button.invoke()

        self.assertEqual(str(card.register_button.cget("state")), "normal")
        self.assertTrue(card.is_pending())

    def test_bulk_register_takes_only_the_safe_defaults(self) -> None:
        service = StubService()
        window = self.open_window(service)
        analysis = make_analysis(1, ["명단 제출", "조건부 신청"])
        analysis.work_items[1].applicability = "conditional"
        analysis.work_items[1].selected = False
        window.present_analysis(analysis)
        window.set_queue_status(1)

        self.assertEqual(len(window._bulk_candidates()), 1)
        window.save_to_outlook()

        self.assertTrue(window.cards[0].item.saved)
        self.assertFalse(
            window.cards[1].item.saved, "조건부 항목은 카드에서 직접 눌러야 한다"
        )
        window.cards[1].register_button.invoke()
        self.assertTrue(window.cards[1].item.saved)


if __name__ == "__main__":
    unittest.main()
