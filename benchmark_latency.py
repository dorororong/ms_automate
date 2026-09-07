"""Opt-in API benchmark: synthetic messages, no Outlook writes or private text logs.

python benchmark_latency.py --repeats 2 --baseline eb0f898
"""
import argparse
import hashlib
import json
import statistics
import subprocess
import time
import types
from pathlib import Path
from unittest.mock import patch

import openai
import upstage_classifier as current

CASES = [
    ("deadline", "9월 10일 오후 5시까지 자료를 제출해 주세요.", 1),
    ("recruitment", "토익 시험 감독관을 선착순으로 모집합니다. 희망자는 신청하세요.\n"
     "일시: 2026. 8. 23(일) 08:30\n수당 8만원, 모집 12명. 당일 오전 10시까지 접수 후 마감합니다.", 1),
    ("multi", "9월 10일 15:00 교무실 회의에 참석하세요.\n"
     "9월 11일까지 설문 결과를 제출하세요.\n매주 화,목 학생 분리수거를 지도하세요.", 3),
    ("completed", "요청하신 자료는 모두 제출 완료했습니다. 추가로 하실 일은 없습니다.", 0),
]


def semantic_match(case, items):
    if case == "completed":
        return not items
    if case == "deadline":
        return (len(items) == 1 and items[0].type == "todo"
                and items[0].due == "2026-09-10" and items[0].due_time == "17:00")
    if case == "recruitment":
        return (len(items) == 1 and items[0].type == "todo" and items[0].intent == "apply"
                and items[0].applicability == "conditional" and not items[0].selected
                and items[0].due is None and items[0].due_time == "10:00"
                and any(c.role == "event_context" for c in items[0].temporal_context))
    return (len(items) == 3 and any(i.start == "2026-09-10T15:00:00" for i in items)
            and any(i.due == "2026-09-11" for i in items)
            and any(i.repeat_freq == "weekly" for i in items))


def main():
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument("--repeats", type=int, default=2)
    parser.add_argument("--baseline", default="eb0f898")
    parser.add_argument("--compact-only", action="store_true")
    parser.add_argument("--output", help="Optional JSON metrics artifact; contains no private source messages")
    args = parser.parse_args()
    code = subprocess.check_output(
        ["git", "show", f"{args.baseline}:upstage_classifier.py"], encoding="utf-8")
    baseline = types.ModuleType("baseline_classifier")
    baseline.__file__ = str(Path(current.__file__).resolve())
    exec(compile(code, "baseline_classifier", "exec"), baseline.__dict__)
    real_client = openai.OpenAI
    results = []
    for repetition in range(args.repeats):
        for case, message, expected in CASES:
            order = [("baseline", baseline), ("compact", current)]
            if args.compact_only:
                order = [("compact", current)]
            if repetition % 2:
                order.reverse()
            for label, module in order:
                calls = []
                clients = []

                def instrumented_client(**kwargs):
                    client = real_client(**kwargs)
                    clients.append(client)
                    real_create = client.chat.completions.create

                    def create(**request):
                        entry = dict(input_chars=sum(len(m["content"]) for m in request["messages"]),
                                     format_chars=len(json.dumps(request["response_format"], ensure_ascii=False)))
                        calls.append(entry)
                        start = time.perf_counter()
                        try:
                            response = real_create(**request)
                            entry.update(usage=response.usage.model_dump() if response.usage else None,
                                         output_chars=len(response.choices[0].message.content or ""))
                            if label == "compact":
                                try:
                                    module._parse_api_response(response, message)
                                except Exception as exc:
                                    entry["validation_error"] = str(exc)
                            return response
                        finally:
                            entry["seconds"] = round(time.perf_counter() - start, 3)
                    client.chat.completions.create = create
                    return client

                start = time.perf_counter()
                row = dict(repetition=repetition + 1, case=case, variant=label)
                try:
                    with patch.object(openai, "OpenAI", instrumented_client):
                        result = module.classify_with_upstage(message, reference_date="2026-08-19")
                    row.update(ok=True, items=len(result.work_items), count_match=len(result.work_items) == expected,
                               semantic_match=semantic_match(case, result.work_items),
                               blocked=sum(bool(i.blocking_review_issues) for i in result.work_items),
                               fields=[dict(type=i.type, intent=i.intent, at=i.start, due=i.due,
                                            due_time=i.due_time, applicability=i.applicability,
                                            context_roles=[c.role for c in i.temporal_context])
                                       for i in result.work_items])
                except Exception as exc:
                    row.update(ok=False, error=type(exc).__name__)
                finally:
                    for client in clients:
                        client.close()
                row.update(seconds=round(time.perf_counter() - start, 3), calls=calls)
                results.append(row)
                print(json.dumps(row, ensure_ascii=False), flush=True)
    summaries = []
    for label in ("baseline", "compact"):
        rows = [r for r in results if r["variant"] == label]
        if not rows:
            continue
        times = [r["seconds"] for r in rows]
        summary = dict(summary=label, n=len(rows), success=sum(r["ok"] for r in rows),
                              mean=round(statistics.mean(times), 3), median=round(statistics.median(times), 3),
                              semantic_pass=sum(r.get("semantic_match", False) for r in rows),
                              under_5=sum(r["ok"] and r["seconds"] <= 5 for r in rows))
        summaries.append(summary)
        print(json.dumps(summary, ensure_ascii=False))
    if args.output:
        from compact_extraction import RULES, SCHEMA
        artifact = dict(baseline=args.baseline, rules_sha256=hashlib.sha256(RULES.encode()).hexdigest(),
                        schema_sha256=hashlib.sha256(json.dumps(SCHEMA, sort_keys=True).encode()).hexdigest(),
                        scope="synthetic API calls; no queue/UI/Outlook timings", results=results, summaries=summaries)
        Path(args.output).write_text(json.dumps(artifact, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
