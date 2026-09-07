# 학교 업무 → Outlook 일정 정리

[구현 리포트 및 상세 사용법](IMPLEMENTATION_REPORT.md)

[GitHub 저장소](https://github.com/dorororong/ms_automate) · 비공개

Windows 데스크톱 앱입니다. 실행하면 화면 오른쪽 아래에 CATMOA식 입력 창이 뜹니다.
다른 프로그램(Outlook 메일, 브라우저, 한글 등)에서 일정·할 일이 적힌 텍스트를
선택해 **Ctrl+Shift+X**로 가져오거나, 메시지를 Ctrl+V로 붙여넣거나, 문서·이미지를
고양이 영역에 드래그하면 Solar Pro 4가 Calendar/Todo 후보를 뽑습니다.

```text
텍스트 선택 / 붙여넣기 / 파일 드롭
  → 입력 내용 읽기
  → 분석 대기열에 순서대로 등록
  → 백그라운드 Solar Pro 4 JSON 분류 (키 없으면 오프라인 규칙)
  → Outlook 기존 일정 시간 겹침·Todo 중복 제외
  → 왼쪽 원문 / 오른쪽 저장 요약 확인
  → 필요한 카드만 `수정`해서 상세 편집
  → `등록` 또는 `미등록` 한 번으로 다음 대기 항목 이동
  → 클래식 Outlook Calendar/Task 기록 (COM) + SQLite 기록
```

메일을 보내거나 Microsoft Graph를 호출하지 않습니다. Outlook은 로컬 COM으로만
조작합니다.

기존 Outlook 항목 조회·수정·삭제는 현재 신규 등록 파이프라인의 범위에 포함하지
않습니다. 해당 기능은 후속 작업으로 보류합니다.

## 실행 환경

- Windows
- Python 3.11 이상
- 클래식 Outlook 데스크톱 앱 (로그인된 상태)
- Upstage API 키 (선택 사항 — 없으면 오프라인 규칙 분석만 사용)

```powershell
python -m pip install -r requirements.txt
python app.py
```

API 키는 환경변수나 프로젝트 루트 `.env`로 설정합니다.

```powershell
$env:UPSTAGE_API = "up_..."
```

```text
UPSTAGE_API=up_...
UPSTAGE_MODEL=solar-pro4
```

`UPSTAGE_API_KEY`, `SOLAR_API_KEY`라는 이름도 인식하며 우선순위는
`UPSTAGE_API` → `UPSTAGE_API_KEY` → `SOLAR_API_KEY`입니다. `.env`는 로컬
전용이며 Git에 커밋하지 않습니다.

## 사용 순서

1. `python app.py`를 실행합니다. 화면 오른쪽 아래에 작은 파란 입력창이 뜹니다.
2. 다음 중 하나로 메시지를 넣습니다.
   - 다른 앱에서 텍스트를 선택하고 **Ctrl+Shift+X**
   - 메시지를 복사한 뒤 입력창에서 **Ctrl+V**
   - PDF/HWP/HWPX/DOCX/TXT/이미지를 고양이 영역에 파일 드롭
3. `일정 찾기`를 누르거나 입력창에서 **Ctrl+Enter**를 누릅니다.
4. 분석이 끝난 뒤 검토창에서 왼쪽 원문과 오른쪽 `저장할 내용`을 나란히 확인합니다.
5. 대부분의 경우 카드의 체크 상태를 확인하고 `등록`을 누릅니다. 세부 내용을 바꿀 때만 카드의 `수정`을 누릅니다.
6. 등록하지 않을 내용은 `미등록`을 누릅니다. 두 버튼 모두 자동으로 다음 대기 항목을 보여줍니다.
7. 분석 중에 새 텍스트·파일을 넣어도 입력을 막지 않습니다. 먼저 들어온 순서대로 처리하고 검토창의 `대기 n건`에 표시합니다.

검토창은 왼쪽에 수정 잠금 상태의 원문, 오른쪽에 저장 요약을 보여줍니다. 원문 수정은
`원문 수정`을 눌렀을 때만 가능하고, 카드 상세 편집도 `수정`을 눌렀을 때만 열립니다.
기준일·재분석·JSON 복사·저장은 오른쪽 위 `⋯` 메뉴에 있습니다. 검토 중에도 오른쪽
아래 입력창은 계속 사용할 수 있습니다.

중복 규칙은 Calendar의 시간 구간 겹침(종료 시간이 없으면 60분으로 간주)과
Todo의 제목+마감일 일치입니다. 같은 메시지에서 생성된 Calendar 시간 겹침과
동일 Todo도 함께 제외하며, 제외 항목은 `참조/제외 n건 보기`에 접어 사유를 보여줍니다.
수정이 필요하면 원문을 다시 넣어 재분석합니다.

작은 창의 버튼:

| 버튼 | 동작 |
|---|---|
| 일정 찾기 | 입력창 내용을 바로 분석 |
| 파일 | 파일 드롭과 같은 파일 읽기 흐름 실행 |
| `⋯` | 선택 텍스트 가져오기, Outlook 연결 확인, 단축키 안내 |

현재 단계의 메인 UI는 신규 등록에 집중합니다. 기존 Outlook 내역 조회·수정·삭제
요청 UI는 별도 관리 영역으로 보류되어 있습니다.

## 파일 입력과 범위

사용자가 직접 선택하거나 드롭한 PDF, HWP, HWPX, DOCX, TXT, Markdown, CSV, TSV 및
이미지를 입력으로 읽습니다. PDF/HWP 계열은 로컬에서 텍스트를 추출한 뒤 분류하며,
이미지는 `pytesseract`와 Windows Tesseract 실행 파일·한국어 언어팩이 설치된 경우에만
읽습니다. OCR이 준비되지 않은 이미지에는 텍스트 붙여넣기를 안내합니다.

쿨메신저 UDB 자동 탐색과 개인정보 탐지/마스킹은 현재 파이프라인에 포함하지 않습니다.
파일을 읽은 뒤 원문은 기존 텍스트 입력과 동일하게 `inputs.original_text`에 저장됩니다.

## 텍스트를 가져오는 방식

Windows에는 "다른 앱에서 지금 선택된 텍스트"를 직접 읽는 API가 없습니다.
그래서 단축키가 눌리면 다음 순서로 동작합니다.

1. 현재 클립보드 내용을 기억합니다.
2. 눌려 있는 Ctrl/Shift/X를 떼고 Ctrl+C를 합성해 보냅니다.
3. 클립보드가 바뀌면 그 텍스트를 가져옵니다.
4. 원래 클립보드 내용을 되돌려 놓습니다.

선택한 것이 없으면 현재 클립보드 내용을 대신 사용합니다. 둘 다 없으면
작은 창에 안내가 표시됩니다.

관리자 권한으로 실행된 프로그램에서는 Windows의 UIPI 정책 때문에 합성 Ctrl+C가
전달되지 않을 수 있습니다. 그때는 직접 Ctrl+C로 복사한 뒤 단축키를 누르면
클립보드 내용으로 처리됩니다.

## 입력 예시

```text
2026-09-10 14:00 학부모 상담 장소: 상담실 알림 15분 전
2026-09-12 운동회 행사 장소: 운동장
2026-09-15 학생 명단 제출
```

시간이 있는 회의·상담은 Outlook 일정으로, 제출·확인·작성 등은 Outlook 작업으로
분류합니다. 날짜만 있는 Calendar 후보는 임의로 00:00/종일로 바꾸지 않고 확인 필요로
남깁니다.

모집·희망자·신청 공지는 날짜마다 나누지 않고 실제 사용자 행동 단위로 묶습니다.
예를 들어 `토익 시험 감독관 모집`, `시험일 2026-08-23 08:30`, `당일 오전 10시까지
접수`는 `토익 시험 감독관 신청` Todo 1건으로 남습니다. 신청은 조건부라 기본 선택이
해제되고, 시험일은 관련 행사 정보로, 날짜가 확정되지 않은 10시는 마감 시각 문맥으로
보존됩니다. 실제 시험 참석 일정이나 `접수 마감 확인` 작업은 자동 생성하지 않습니다.

`key=value` 형식 입력도 인식합니다.

```text
title=학부모 상담
target=calendar
start=2026-09-10 14:00
duration_minutes=30
location=상담실
reminder_minutes=15
memo=상담 자료 준비
```

`target=both`를 쓰면 같은 입력에서 일정과 작업을 함께 검토할 수 있습니다.

## 분류 체계

**1차 (scope) — LLM이 판단**

| 분류 | 기준 | Outlook |
|---|---|---|
| 담임 | 최종 대상이 **학생**. 전달·안내·배부·지도 | 등록함 |
| 업무 | 최종 대상이 **나(교사)**. 제출·참석·이수·신청 | 등록함 |
| 참조 | 읽고 넘기면 되는 안내 | **등록 안 함** (SQLite에만 기록) |

**2차 (action) — 코드가 필드에서 파생. LLM에게 묻지 않음**

| 조건 | action | Outlook |
|---|---|---|
| 마감(`due`) 있음 | `todo_with_due_date` | 작업 + 마감일 + 내역 |
| 일시(`at`) 있음 | `todo_with_date` | 캘린더 + 알림 |
| 둘 다 없음 | `todo` | 작업 (마감 없음) |
| scope=참조 | `none` | 등록 안 함 |

LLM에게 2차 라벨까지 물었을 때 실측 7%가 자기모순(`todo_with_due_date` 인데
`due=null`)이었습니다. 사실만 받아 파생하면 그 모순이 구조적으로 불가능합니다.
확인 창에서 분류나 일시/마감을 고치면 라벨이 즉시 다시 파생됩니다.

**반복** — `repeat_freq`(daily/weekly/monthly) + `repeat_detail`("화,목")로
추출해 Outlook 반복 작업·일정으로 등록합니다.

추출 결과에는 실제 행동을 나타내는 `intent`, 조건부 여부를 나타내는
`applicability`, 완료/안내 상태인 `source_state`, `condition`, `due_time`,
`temporal_context`, `checklist`, `evidence`, `review_issues`가 함께 보존됩니다.
날짜가 여러 개여도 `event_context`·`external_deadline`·`historical` 같은 역할로
분리되어 주 행동의 등록 대상과 섞이지 않습니다.

## 검증 결과와 현재 품질

현재 기본 프롬프트는 `semantic-v4-compact`입니다. API 필수 필드를 22개에서
12개로 줄였고, 빈 선택 필드·내부 ID·기본값은 앱에서 채웁니다. 관련 날짜는
역할과 원문만 받으며 기존 날짜·조건·중복 검사를 거친 뒤 검토창에 표시합니다.
상세 속도 비교는 [LATENCY_REPORT.md](LATENCY_REPORT.md)를 참고하세요.
아래 32건 수치는 간소화 **이전** 버전의 결과이며 현재 버전의 정확도 점수가 아닙니다.

회귀 fixture인 `test_set_30.xlsx`의 `테스트셋` 시트 본문 32건을 Solar Pro 4에
발송일과 함께 전달해 확인했습니다. 이 평가는 신규 메시지 전체의 일반화 점수가
아니라 현재 프롬프트와 의미 필드의 기준선입니다.

- 결정적 로컬 테스트: `python selftest.py --offline` → 27/27 통과
- API 포함 자가 점검: `python selftest.py` → 30/30 통과
- 문법·의존성 검사: `compileall` 통과, `pip check` 이상 없음
- 개선 후 마지막 Solar Pro 4 32건: 31/32건 응답 확보, 1건은 2회 시간 초과 후 fallback 대상
- 마지막 성공 응답시간: 평균 31.92초, 중앙값 27.62초, 95백분위 92.03초, 최대 97.53초
- 성공 응답 중 60초 이내 29건, 90초 이내 29건, 원문에 없는 `00:00` Calendar 0건
- 변경 전 의미 기준 수동 감사: 완전 통과 12건, 부분 통과 11건, 개선 필요 9건

개선 후 호출은 첫 시도 60초, 재시도 90초, 최대 2회로 실행합니다. 측정에서는 한 번의
반복에서 32/32건, 마지막 반복에서 31/32건이 성공해 서버 응답 변동이 확인됐습니다.
시간 초과나 구조화 응답 오류로 오프라인 규칙으로 전환되면 후보 선택을 해제해 사용자가
명시적으로 선택하도록 했습니다. 구조가 맞지 않는 JSON, 필수 필드 누락, 원문에 없는
수행 시각, 조건 없는 조건부 항목은 등록 전에 차단합니다. 따라서 현재 버전은 결과를
확인한 뒤 등록하는 베타 수준으로 사용해야 합니다.

주요 반복 사례도 확인했습니다. 토익 모집은 3회 모두 `신청 Todo 1건`으로
안정적이었지만, 충전·책상 배열은 3회 중 2회만 2건 분리를 유지했고, 연수·감독은
권장 영상의 조건부 처리와 긴 공지의 행동 단위 분리는 반복에 따라 변동했습니다.
이런 의미 판단은 여전히 Solar Pro 4와 프롬프트의 영향을 받으므로, 후보 수가 많거나
조건이 복잡한 메시지는 카드의 원문과 확인 사유를 함께 확인해야 합니다.

## 사용자 프로필

담임/업무 판정에는 사용자 역할 정보가 필수입니다. 배포 후 `profile.example.json`을
복사해 로컬 전용 `profile.json`을 만들고 수정하세요.

```json
{
  "role": "예: 중학교 교사",
  "department": "",
  "subject": "",
  "notes": ""
}
```

## 기준일

`내일`·`다음주 월요일` 계산 기준입니다. 검토창의 `⋯ → 기준일 설정`에서 바꿀 수
있고 기본값은 오늘입니다. 메시지 발송일을 넣으면 정확해집니다. 본문에 발송일·작성일이
적혀 있으면 모델이 그것을 우선하도록 프롬프트에 규칙이 들어 있습니다.

## 분류 방식

API 키가 있으면 Upstage `solar-pro4`를 OpenAI 호환 endpoint
`https://api.upstage.ai/v1`로 호출합니다. 첫 요청은 60초, 재시도는 90초까지
기다리며 최대 2회 시도합니다. 모두 실패하면 오프라인 규칙 분석으로 넘어가고,
확인 창 아래에 그 사실을 표시합니다. 학교 네트워크에서 외부 API가 막혀 있어도
워크플로우가 멈추지 않습니다.

추론은 `reasoning_effort="none"`이며, 간소화한 JSON Schema와 로컬 형식 검증을
함께 사용합니다. 기존 긴 프롬프트와 스키마는 매번 전송하지 않습니다.
5초는 목표 응답시간이며 제한시간이나 보장 시간이 아닙니다.

실제 API 응답시간은 메시지 길이와 서버 상태에 따라 달라집니다.
분석은 백그라운드 FIFO 대기열에서 실행되므로 입력창은 계속 사용할 수 있지만,
대기 중인 항목이 오프라인 규칙으로 바뀌면 긴 공지의 의미 분리 품질이 낮아질 수
있습니다. 오프라인 전환 결과는 기본 선택을 해제하며, 등록 전에 검토창의 조건·확인
사유를 확인하세요.

시간이 불확실한 Calendar 항목의 `start`는 `null`로 남기며 날짜를 임의로
추측하지 않습니다. 확인 사유가 해소되기 전에는 등록을 차단합니다. 시간 일정의 `end`가
없으면 Outlook 저장 시 60분을 씁니다. Todo는 `due=null`이어도 등록할 수 있지만
`due_time`은 설명에 보존합니다.

## JSON 출력 형식

확인 창의 `JSON 복사` / `JSON 저장`이 내보내는 구조입니다.

```json
{
  "schema_version": 2,
  "input_id": 1,
  "source": "upstage",
  "model": "solar-pro4",
  "items": [
    {
      "type": "calendar",
      "title": "체험학습 관련 회의",
      "start": "2026-09-05T15:00:00",
      "end": null,
      "due": null,
      "due_time": null,
      "intent": "attend",
      "applicability": "required",
      "condition": null,
      "source_state": "pending",
      "description": "관련 회의 참석",
      "location": "상담실",
      "category": null,
      "all_day": false,
      "reminder_minutes": 15,
      "reminder": null,
      "saved": false,
      "outlook_entry_id": null,
      "temporal_context": [],
      "checklist": [],
      "evidence": [],
      "review_issues": [],
      "group_id": null
    }
  ],
  "ignored": []
}
```

## SQLite

DB 파일은 `data/work_items.sqlite3`입니다.

- `inputs`: 원문, 생성 시간, 기준일·발송일·시간대 등 `context_json`
- `work_items`: 분류, 제목, 시간/마감일, 장소/범주/알림, 설명, `saved`,
  Outlook `EntryID`, 오류, `due_time`, 의미 필드 `semantic_json`

CATMOA와 같은 방식으로 분석 직후 pending 상태로 먼저 기록하고, Outlook 기록에
성공하면 `saved=1`과 `outlook_entry_id`를, 실패하면 `last_error`를 남깁니다.
성공한 항목은 체크박스가 비활성화되어 같은 항목을 두 번 등록하지 않습니다.

기존 Outlook 항목의 조회·수정·삭제는 별도 관리 파이프라인으로 설계할 영역이며,
현재 메인 UI에서는 연결하지 않았습니다. 현재 등록 파이프라인은 새 Calendar/Task를
등록하고 그 결과를 SQLite에 기록하는 데 집중합니다.

## Outlook 연결 전제

`pywin32`로 `Outlook.Application` COM을 사용하므로 클래식 Outlook이 설치되어
있어야 합니다. New Outlook만 설치된 환경에서는 동작하지 않습니다. Outlook에
로그인한 상태로 실행하세요. `Outlook 연결 확인`은 연결만 검사하며 새 항목을
만들지 않습니다.

Calendar/작업 읽기와 기록은 현재 PC의 Outlook 프로필에 대한 로컬 COM
자동화입니다. Microsoft Graph 같은 외부 HTTP API를 쓰지 않으므로 학교
네트워크의 API 차단과 분리되어 있습니다. 같은 Microsoft 계정의 Outlook 작업은
Microsoft To Do와 동기화될 수 있습니다.

## 파일 구성

| 파일 | 역할 |
|---|---|
| `app.py` | CATMOA식 드롭/붙여넣기 입력 창, 전역 단축키, FIFO 분석 대기열 |
| `file_input.py` | 사용자가 드롭한 문서·이미지의 로컬 텍스트 추출 |
| `windows_drop.py` | Tk 창의 Windows 파일 드래그앤드롭 연결 |
| `hotkey.py` | Windows `RegisterHotKey` 전역 단축키 (전용 스레드 + 메시지 루프) |
| `selection.py` | 선택 텍스트 캡처, 클립보드 읽기/쓰기/복원 |
| `review.py` | 왼쪽 원문·오른쪽 저장 요약, 수정, 등록/미등록, JSON 내보내기 |
| `cards.py` | WorkItem 한 건을 편집하는 카드 위젯 |
| `service.py` | 분류 → 중복 검사 → SQLite pending → Outlook 기록 → saved/failed |
| `classifier.py` | 오프라인 규칙 분류 |
| `upstage_classifier.py` | Solar Pro 4 구조화 JSON 분류 |
| `compact_extraction.py` | 간소화 프롬프트·API 스키마·로컬 기본값 복원 |
| `benchmark_latency.py` | 합성 메시지 API 속도·핵심 의미 비교 (명시 실행 시 API 호출) |
| `test_compact_extraction.py` | 축약 응답·조건·반복·시간 차단 회귀 검사 |
| `outlook_adapter.py` | 클래식 Outlook COM 어댑터 |
| `storage.py`, `models.py` | SQLite 저장소와 WorkItem 모델 |
| `selftest.py` | Outlook에 쓰지 않는 로컬/연결 검증 스크립트 |
| `profile.example.json` | 사용자 프로필 로컬 설정 예시 (`profile.json`은 Git 제외) |

## 단축키를 바꾸려면

`app.py`의 `VK_X`와 `HOTKEY_LABEL`, 그리고 `HotkeyListener(modifiers=...)`를
바꿉니다. 다른 프로그램이 이미 그 조합을 쓰고 있으면 작은 창에 등록 실패가
표시되고, 그때도 입력창의 붙여넣기·파일 선택으로 같은 작업을 할 수 있습니다.

## 개인정보

원문은 SQLite `inputs.original_text`에 저장됩니다. 현재 요구사항에 따라 PII 탐지·마스킹은
수행하지 않습니다.
민감한 학교 메시지를 다룰 때는 로컬 `.env`와 DB 파일을 공유하지 마세요.
