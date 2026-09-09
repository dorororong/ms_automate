# 작업 트리 (WORK_TREE)

`ms_automate` 저장소의 현재 상태를 한 장으로 보는 문서입니다. 무엇이 끝났고,
무엇이 부분이며, 무엇이 보류인지를 트리로 적었습니다. 작업을 시작하기 전에
해당 가지의 **관련 파일**과 **완료 기준**을 먼저 확인하세요.

- 기준 커밋: `23bddd9` (2026-09-08) + 미커밋 작업 (카드 단위 처리 · 메시지 누적 · Outlook 미러)
- 상태 표기: `[x]` 완료 · `[~]` 동작하지만 품질/범위 미완 · `[ ]` 보류·미착수
- 상세 근거: [README.md](README.md) · [IMPLEMENTATION_REPORT.md](IMPLEMENTATION_REPORT.md) ·
  [LATENCY_REPORT.md](LATENCY_REPORT.md) · [ACTION_EXTRACTION_SPEC.md](ACTION_EXTRACTION_SPEC.md)

---

## 1. 최근 작업 이력

| 커밋 | 날짜 | 내용 | 남긴 것 |
|---|---|---|---|
| `8b06894` | 09-07 23:44 | CATMOA식 Outlook 파이프라인 최초 구현 (23파일, +7001줄) | 입력창·검토창·COM 어댑터·SQLite·문서 일체 |
| `eb0f898` | 09-08 00:37 | Solar 분류 및 등록 가드 강화 | 원문에 없는 시각 차단, 조건부 기본 미선택, `selftest.py` 확장 |
| `4451361` | 09-08 07:46 | Solar 추출 payload 축소 + 지연 실측 | `compact_extraction.py`, `benchmark_latency.py`, 벤치 JSON 3종, `LATENCY_REPORT.md` |
| `23bddd9` | 09-08 07:59 | 검토 카드 편집 속도 개선 | 카드에서 제목·일시/마감 직접 편집, 등록 가능 여부를 편집값 기준으로 재평가 |
| (미커밋) | 09-09 | 카드 단위 처리 · 메시지 누적 · Outlook 로컬 미러 | `outlook_mirror.py`, `test_review_flow.py`, 체크박스 제거, `outlook_mirror`/`sync_state` 표 |
| (미커밋) | 09-09 | 상태만 남긴 작은 창 · 트레이 아이콘 | `tray.py`, 입력창·실행 버튼 제거, 항상 위, 일시중지 |
| (미커밋) | 09-09 | 마감일 캘린더 표시 · 카드 가독성 · 창 폭 축소 | `outlook_mirror` 표시 행, 제목 줄바꿈, 등록/미등록 간격, `120×160` |
| (미커밋) | 09-09 | 앱 아이콘 · README 재작성 | `icons.py`, 파이썬 기본 아이콘 교체, README 전면 정리 |

흐름 요약: **파이프라인 구축 → 잘못된 등록 차단 → API 지연 절감 → 검토 UI 손질 →
검토 조작 자체를 없애기**. 기능은 한 번에 만들어졌고, 이후 작업은 모두
*정확도 · 속도 · 검토 비용*을 깎는 쪽입니다.

미커밋 작업의 핵심은 두 가지입니다. 체크박스로 고른 뒤 아래에서 한꺼번에 저장하던
흐름을 카드의 `등록`/`미등록` 버튼으로 바꿨고, 검토 중 분석이 끝난 메시지를 기다리게
하지 않고 같은 창에 이어 붙입니다. 카드마다 등록하면 클릭할 때마다 Outlook 폴더 전체를
COM으로 훑어야 해서, 중복 검사가 읽을 로컬 미러(`outlook_mirror`)를 만들었습니다.

`4451361`에서 평균 응답이 12.178초 → 4.606초로 줄었지만, 완료 안내 2건을 뺀
실제 추출 6회 평균은 5.595초로 목표 5초는 아직 미달입니다 (LATENCY_REPORT.md).

---

## 2. 실행 흐름 트리

```text
app.py (MiniWindow · 오른쪽 아래 350x160 · 항상 위 · 상태 표시 전용)
├── 입력 (창에는 입력 위젯이 없다)
│   ├── Ctrl+Shift+X 선택 텍스트        → hotkey.py + selection.py
│   ├── Ctrl+V 클립보드                 → app.paste_from_clipboard
│   ├── 파일 드롭 · 창 클릭              → windows_drop.py + file_input.py
│   └── 트레이 일시중지 중에는 받지 않음  → tray.py
│                                          (PDF·HWP·HWPX·DOCX·TXT·MD·CSV·TSV·이미지 OCR)
├── FIFO 분석 대기열 (백그라운드 스레드)   → app.py AnalysisRequest / queue
│   └── service.WorkflowService
│       ├── upstage_classifier.classify_with_upstage   (키 있음: solar-pro4)
│       │   └── compact_extraction.RULES / SCHEMA / expand
│       │       └── 실패·타임아웃 2회 → 아래 오프라인 경로, 후보 기본 선택 해제
│       └── classifier.classify_text                   (오프라인 규칙)
├── 저장 및 대조
│   ├── storage.Database → inputs / work_items 에 pending 기록
│   └── outlook_mirror   → 로컬 사본에서 시간 겹침 · Todo 제목+마감 중복 제외
├── 검토 (review.ReviewWindow 760x620)
│   ├── 왼쪽: 초점 메시지의 원문 (기본 잠금, `원문 수정`으로 해제)
│   ├── 오른쪽: MessageGroup 목록 — 분석이 끝난 메시지가 계속 이어 붙음
│   │   └── cards.WorkItemCard — 제목·일시/마감 즉시 편집 + `등록`/`미등록` 버튼
│   ├── `⋯` 메뉴: 기준일 · 재분석 · JSON 복사/저장
│   └── 접힘 영역: `참조/제외 n건 보기`
└── 카드별 등록 / 미등록 (또는 `모두 등록`)
    ├── outlook_adapter (클래식 Outlook COM) → storage 에 saved + EntryID / failed
    └── outlook_mirror.write_through → 다음 카드의 중복 검사에 즉시 반영
```

## 3. 모듈 의존 트리

```text
app.py
├── selection.py        클립보드 캡처·복원 (프로젝트 내부 의존 없음)
├── hotkey.py           RegisterHotKey 전용 스레드 (내부 의존 없음)
├── tray.py             알림 영역 아이콘 전용 스레드 → icons.py
├── icons.py            PNG → ICO 변환, 창·작업 표시줄 아이콘 (내부 의존 없음)
├── windows_drop.py     파일 드래그앤드롭 연결 (내부 의존 없음)
├── file_input.py       로컬 텍스트 추출 (내부 의존 없음)
├── theme.py            공통 스타일 (내부 의존 없음)
├── review.py
│   ├── cards.py → models.py, theme.py
│   ├── service.py
│   └── outlook_adapter.py
└── service.py
    ├── classifier.py            → models.py
    ├── upstage_classifier.py    → classifier.py, models.py, compact_extraction.py
    ├── outlook_adapter.py       → models.py
    ├── outlook_mirror.py        → outlook_adapter.py, storage.py, models.py
    └── storage.py               → models.py

models.py             리프. WorkItem · scope/action 파생 규칙 · 등록 가능 판정
compact_extraction.py 리프. 프롬프트 규칙 + API 스키마 + 축약 응답 복원

검증·계측 (app.py 실행 그래프 밖)
├── selftest.py                  --offline 32/32, API 포함 30/30
├── test_compact_extraction.py   축약 계약 회귀 7/7
├── test_review_flow.py          검토창 Tk 위젯 회귀 8/8
└── benchmark_latency.py         API 호출 실측 → latency_benchmark*.json
```

의존 방향이 단방향(UI → service → 분류/저장/COM → models)이라
**`models.py`를 고치면 거의 전부가 영향받고, `cards.py`/`review.py`를 고치면 UI 안에서 끝납니다.**

---

## 4. 기능별 작업 트리

```text
[x] 1. 입력 수집
    [x] 1.1 전역 단축키 Ctrl+Shift+X            hotkey.py, selection.py
    [x] 1.2 붙여넣기 / 직접 입력                 app.py
    [x] 1.3 파일 드롭·선택                       windows_drop.py, file_input.py
    [~] 1.4 이미지 OCR                           file_input.py
        └─ Tesseract 실행 파일 + 한국어 언어팩이 설치된 PC에서만 동작
    [ ] 1.5 쿨메신저 UDB 자동 탐색               범위 밖 (미포함)

[~] 2. 의미 추출
    [x] 2.1 Solar Pro 4 구조화 JSON 호출         upstage_classifier.py
    [x] 2.2 축약 프롬프트/스키마 semantic-v4-compact   compact_extraction.py
    [x] 2.3 오프라인 규칙 fallback               classifier.py
    [x] 2.4 scope(LLM) → action(코드 파생) 분리   models.py
    [~] 2.5 긴 공지의 행동 단위 분리              반복 실행 시 결과 변동 (README 검증 결과)
    [~] 2.6 조건부·완료 안내 처리                 조건부는 기본 미선택으로 방어 중

[x] 3. 잘못된 등록 차단 (eb0f898)
    [x] 3.1 원문에 없는 시각의 Calendar 차단      upstage_classifier.py
    [x] 3.2 필수 필드 누락·스키마 불일치 차단      upstage_classifier.py, compact_extraction.py
    [x] 3.3 조건 없는 조건부 항목 차단            models.py
    [x] 3.4 Outlook 중복 조회 실패 시 등록 보류    service.py

[x] 4. 중복·겹침 제외
    [x] 4.1 Calendar 시간 구간 겹침 (end 없으면 60분)   service.py
    [x] 4.2 Todo 제목+마감 일치                   service.py
    [x] 4.3 같은 메시지 안에서 생성된 중복          service.py
    [x] 4.4 제외 사유 표시 (`참조/제외 n건 보기`)   review.py
    [x] 4.5 검사 대상을 로컬 미러로 전환            service.known_entries

[x] 4.6 한 메시지 안의 여러 건도 각자 처리      test_review_flow.py

[~] 5. 검토 UI
    [x] 5.1 왼쪽 원문 / 오른쪽 저장 요약           review.py
    [x] 5.2 카드에서 제목·일시/마감 즉시 편집       cards.py (23bddd9)
    [x] 5.3 편집값 기준 등록 가능 재평가            cards.can_register_now (23bddd9)
    [x] 5.4 카드별 `등록`/`미등록` 즉시 처리         cards.py, review._register_card
    [x] 5.5 `미등록` 되돌리기                       cards._on_dismiss_clicked
    [x] 5.6 안전 기본값만 담는 `모두 등록`           review.save_to_outlook
    [x] 5.7 분석 끝난 메시지를 같은 창에 누적        review.MessageGroup, app._show_next_analysis
    [x] 5.8 초점·왼쪽 원문이 미처리 메시지를 따라감   review._advance_focus
    [x] 5.9 남은 카드·분석이 없을 때만 창 종료       review._after_action
    [ ] 5.10 기존 Outlook 항목 조회·수정·삭제 UI    보류. 모듈은 _archive 로 분리됨
    [x] 5.11 검토창을 작은 창과 겹치지 않게 배치      review._placement
    [x] 5.12 긴 제목 줄바꿈 · 끝까지 보이기           cards._fit_title_height
    [x] 5.13 등록/미등록 오클릭 방지 간격             cards.py actions

[x] 10. 작은 창과 트레이
    [x] 10.1 입력창·실행 버튼 제거, 상태 한 줄만      app._build_ui
    [x] 10.2 세로 절반 (320 → 160)                  app.WINDOW_HEIGHT
    [x] 10.3 켜져 있는 동안 항상 위                   app.__init__
    [x] 10.4 상태 문구 한 곳에서 생성                 app._refresh_status
    [x] 10.5 트레이 아이콘 · 창 보이기/일시중지/종료   tray.py
    [x] 10.6 X 는 트레이로 내리기                     app._on_close_button
    [x] 10.7 일시중지 중 입력 차단                    app._enqueue_analysis
    [x] 10.8 가로 폭 1/3 (350 → 120)                 app.WINDOW_WIDTH
    [x] 10.9 짧은 상태 + 자세한 설명은 `⋯` 메뉴       app._refresh_status
    [x] 10.10 트레이 툴팁은 이름만                    app._sync_tray_tooltip
    [x] 10.11 앱 아이콘 (창·작업 표시줄·트레이)        icons.py, tray_icon.png

[x] 6. 저장·기록
    [x] 6.1 pending → saved/failed 상태 기록       storage.py, service.py
    [x] 6.2 EntryID · last_error 보존              storage.py
    [x] 6.3 semantic_json 의미 필드 보존           models.py, storage.py
    [x] 6.4 JSON 복사/저장 (schema_version 2)      review.py

[x] 7. Outlook 연동
    [x] 7.1 클래식 COM Calendar/Task 생성          outlook_adapter.py
    [x] 7.2 반복 일정/작업 (repeat_freq/detail)     outlook_adapter.py
    [x] 7.4 마감일 종일 표시 생성                   outlook_adapter.save_due_marker
    [x] 7.5 표시는 겹침 검사에서 제외               service._is_due_marker
    [x] 7.6 표시 on/off 토글                        cards.py `상세`
    [ ] 7.3 New Outlook / Microsoft Graph 지원      미지원. 로컬 COM 전제

[~] 8. 검증·계측
    [x] 8.1 오프라인 자가 점검 32/32               selftest.py --offline
    [x] 8.2 API 포함 점검 30/30                    selftest.py (미러 추가 전 측정)
    [x] 8.3 축약 계약 회귀 7/7                     test_compact_extraction.py
    [x] 8.4 검토창 회귀 8/8                        test_review_flow.py
    [x] 8.5 지연 실측                              benchmark_latency.py
    [~] 8.6 32건 의미 회귀                         현재(compact) 버전으로 재채점 안 됨
    [ ] 8.7 PII 탐지·마스킹                        요구사항상 미수행

[~] 9. Outlook 로컬 미러
    [x] 9.1 outlook_mirror / sync_state 표         storage.py
    [x] 9.2 앱 시작 시 전체 동기화                  outlook_mirror.OutlookMirror.start
    [x] 9.3 5분 주기 재동기화 (전용 COM 스레드)      outlook_mirror._worker
    [x] 9.4 등록 시 write-through                  service.record → mirror.write_through
    [x] 9.5 사라진 항목 제거                        storage.sync_outlook_mirror
    [x] 9.6 동기화 상태 표시                        app._on_mirror_synced / _on_mirror_error
    [~] 9.7 외부 프로그램이 만든 항목의 사각          다음 동기화까지 최대 5분
    [ ] 9.8 증분 동기화 (LastModificationTime)      미적용. 매번 전체 재조회
```

---

## 5. 다음 작업 후보

우선순위는 "현재 릴리즈 판정 = 검토형 베타"를 올리는 데 필요한 순서입니다.

### P1 — 현재 프롬프트로 32건 재채점
- **왜**: README·IMPLEMENTATION_REPORT의 32건 수치(완전 12 / 부분 11 / 개선 필요 9)는
  모두 축약 이전 기준선입니다. `4451361`에서 필수 필드를 22 → 12개로 줄인 뒤의
  의미 정확도는 측정되지 않았고, 지금은 **속도 개선만 확인**된 상태입니다.
- **관련**: `test_set_30.xlsx`(`테스트셋` 시트), `upstage_classifier.py`, `compact_extraction.py`
- **완료 기준**: 축약 버전의 32건 채점표가 리포트에 들어가고, 문서에서 "간소화 전 기준선"
  단서가 사라진다. API 비용이 발생하므로 실행 전에 합의할 것.

### P2 — 반복 실행 편차가 큰 사례 고정
- **왜**: 충전·책상 배열은 3회 중 2회만 2건 분리를 유지했고, 연수·감독의 조건부 처리는
  회차마다 변동했습니다. 토익 모집만 3회 모두 안정적이었습니다.
- **관련**: `compact_extraction.py`(RULES), `ACTION_EXTRACTION_SPEC.md`
- **완료 기준**: 같은 입력 3회 반복에서 항목 수와 조건부 여부가 동일.

### P3 — 5초 목표의 결론 내기
- **왜**: 실제 추출 6회 평균 5.595초, 5초 이내 2/6. 출력 축소만으로는 한계이며,
  대기 시간과 생성 시간을 분리 계측하지 않아 지연의 내부 원인이 미확정입니다.
- **관련**: `benchmark_latency.py`, `upstage_classifier.py`
- **완료 기준**: 첫 응답까지의 대기와 전체 생성 시간을 나눠 측정한 뒤,
  5초를 목표로 유지할지 목표치를 조정할지 문서에 명시한다.

### P4 — 벤치마크 산출물 정리
- **왜**: `latency_benchmark.json` / `_final.json` / `_v4.json` 세 개가 루트에 있고,
  `_final`은 최종 제품 버전이 아니라 중간 10필드 실험이라 파일명이 오해를 부릅니다.
- **완료 기준**: 결과 폴더로 옮기거나 실험 순서대로 이름을 바꾼다.

### P5 — 미러 동기화 비용 줄이기
- **왜**: 지금은 5분마다 Calendar/작업 폴더 전체를 다시 읽습니다. 이 PC 프로필(13건)은
  1.4초였지만 항목이 수천 건인 프로필에서는 주기마다 비용이 커집니다.
- **관련**: `outlook_mirror.py`(`sync_once`), `outlook_adapter.read_entries`
- **완료 기준**: `Items.Restrict("[LastModificationTime] > …")`로 변경분만 읽고,
  삭제 반영을 위한 전체 재조회는 더 긴 주기로 분리한다. 큰 프로필에서 주기당
  소요 시간을 실측해 기록한다.

### P6 — 보류 가지 (착수 전 범위 결정 필요)
- 기존 Outlook 항목 조회·수정·삭제 UI. 모듈(`command_view.py`, `outlook_command.py`,
  `outlook_view.py`)은 `<project-parent>/_archive/ms_automate_unused_20260907` 에 보관 중.
- PII 탐지·마스킹, 쿨메신저 UDB 자동 탐색.

---

## 6. 검증 명령

오프라인 자가 점검 (Outlook 미러 포함):

```powershell
python selftest.py --offline
```

축약 계약 회귀:

```powershell
python -m unittest test_compact_extraction -v
```

검토창 회귀 (Tk 창을 잠깐 띄웁니다):

```powershell
python -m unittest test_review_flow -v
```

API 포함 점검 (Upstage 호출, 비용 발생):

```powershell
python selftest.py
```

지연 실측 (Upstage 호출, 비용 발생):

```powershell
python benchmark_latency.py --repeats 2 --baseline eb0f898
```

`selftest.py`는 Outlook에 실제 항목을 쓰지 않습니다.

## 7. 문서 지도

| 문서 | 쓰임 |
|---|---|
| `README.md` | 사용자용. 설치·조작·분류 체계·JSON 형식·SQLite 스키마 |
| `IMPLEMENTATION_REPORT.md` | 구현 근거. 파이프라인·중복 규칙·32건 평가·제한사항 |
| `LATENCY_REPORT.md` | 축약 프롬프트 설계와 속도 실측 (2026-09-08) |
| `ACTION_EXTRACTION_SPEC.md` | 행동 단위 추출 규칙 명세 |
| `WORK_TREE.md` | 이 문서. 진행 상태와 다음 작업 |

문서를 고칠 때 이 파일의 상태 표기도 같이 옮기세요. 특히 P1이 끝나면
4장의 `2.5`·`8.6`과 1장의 "기준선" 문장이 함께 바뀝니다.
