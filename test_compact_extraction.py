import unittest
from types import SimpleNamespace
import json

from compact_extraction import expand
from upstage_classifier import _parse_api_response, UpstageAPIError


def payload(**changes):
    item = dict(scope="업무", type="todo", title="자료 제출", intent="submit",
                at=None, due="2026-09-10", due_time=None, condition=None, quote="자료 제출", context=[],
                repeat_freq="none", repeat_detail="")
    item.update(changes)
    return {"items": [item]}


def parse(data, source="9월 10일까지 자료 제출"):
    response = SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(
        content=json.dumps(data, ensure_ascii=False)))])
    return _parse_api_response(response, source).work_items


class CompactTests(unittest.TestCase):
    def test_defaults_and_deadline(self):
        item = parse(payload())[0]
        self.assertEqual(item.due, "2026-09-10")
        self.assertTrue(item.selected)
        self.assertEqual(item.repeat_freq, "none")

    def test_conditional_recruitment(self):
        item = parse(payload(intent="apply", condition="희망자", due=None,
                             due_time="10:00", warning="접수 날짜 확인",
                             context=[dict(role="event_context", raw_text="8월 23일 08:30 시험")]))[0]
        self.assertEqual(item.applicability, "conditional")
        self.assertFalse(item.selected)
        self.assertEqual(item.due_time, "10:00")
        self.assertTrue(item.blocking_review_issues)
        self.assertIn("8월 23일", item.temporal_context[0].raw_text)

    def test_repeat_checklist_preserved(self):
        item = parse(payload(repeat_freq="weekly", repeat_detail="화,목",
                             checklist=["충전", "책상 배열"], location="교실"))[0]
        self.assertEqual(item.repeat_freq, "weekly")
        self.assertEqual(len(item.checklist), 2)
        self.assertEqual(item.location, "교실")

    def test_invented_time_still_blocked(self):
        item = parse(payload(type="calendar", at="2026-09-10T09:00:00", due=None))[0]
        self.assertIsNone(item.start)
        self.assertFalse(item.selected)

    def test_missing_core_rejected(self):
        data = payload()
        del data["items"][0]["condition"]
        with self.assertRaises(UpstageAPIError):
            parse(data)

    def test_invalid_optional_types_rejected(self):
        for changes in [dict(checklist="x"), dict(all_day="false"),
                        dict(reminder_minutes=True), dict(context=[{}]),
                        dict(repeat_freq="sometimes"), dict(quote="")]:
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                expand(payload(**changes))

    def test_completed_message_can_be_empty(self):
        self.assertEqual(parse({"items": []}), [])


if __name__ == "__main__":
    unittest.main()
