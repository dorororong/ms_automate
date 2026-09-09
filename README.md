# 학교 업무 메시지 → Outlook 일정·할 일

> ### 🐈 [CATMOA](https://github.com/progh2/catmoa) 기반
>
> - 교사 개발자 해커톤 결과물인 **CATMOA**에 공동 참여
> - 메시지 → 일정 추출 → 캘린더 등록 흐름과 작은 창 UX를 CATMOA에서 차용
> - 학교 환경의 **Google API 사용 제한** 때문에 Outlook 버전으로 재개발
> - 자세한 내용: [배경](#배경)

학교 공지·메일·메신저 메시지에서 일정과 할 일을 뽑아 클래식 Outlook에 등록하는
Windows 데스크톱 앱.

```text
텍스트 선택(Ctrl+Shift+X) · 붙여넣기(Ctrl+V) · 파일 드롭
        ↓
FIFO 분석 대기열 (백그라운드)
        ↓
Solar Pro 4 구조화 JSON 분석  ─ 키 없음/실패 → 오프라인 규칙 분석
        ↓
SQLite pending 기록
        ↓
로컬 Outlook 미러로 시간 겹침 · 중복 제외
        ↓
검토창: 왼쪽 원문 / 오른쪽 카드
  카드마다 [등록] [미등록] — 한 건씩 즉시 처리
  분석이 끝난 다음 메시지는 같은 창에 이어 붙음
        ↓
클래식 Outlook Calendar / Task 생성 (COM)
  마감 있는 작업은 마감일 종일 표시도 함께
        ↓
SQLite saved + EntryID 기록
```

- 메일 발송 없음. Microsoft Graph 호출 없음. Outlook은 로컬 COM으로만 조작
- 외부 API가 막힌 학교 네트워크에서도 등록 흐름 동작

| 문서 | 내용 |
|---|---|
| [IMPLEMENTATION_REPORT.md](IMPLEMENTATION_REPORT.md) | 구현 근거, 평가 기록 |
| [LATENCY_REPORT.md](LATENCY_REPORT.md) | 프롬프트 축소, 속도 실측 |
| [ACTION_EXTRACTION_SPEC.md](ACTION_EXTRACTION_SPEC.md) | 행동 단위 추출 규칙 |
| [WORK_TREE.md](WORK_TREE.md) | 진행 상태, 다음 작업 |

- 저장소: [dorororong/ms_automate](https://github.com/dorororong/ms_automate) (비공개)
- 기반이 된 원본: [progh2/catmoa](https://github.com/progh2/catmoa) — 변경 가하지 않음

---

## 목차

1. [배경](#배경)
2. [빠른 시작](#빠른-시작)
3. [화면](#화면)
4. [입력 방법](#입력-방법)
5. [분석](#분석)
6. [Outlook 등록](#outlook-등록)
7. [저장 데이터](#저장-데이터)
8. [검증 결과](#검증-결과)
9. [개발자용](#개발자용)
10. [문제 해결](#문제-해결)
11. [개인정보](#개인정보)
12. [범위 밖](#범위-밖)

---

## 배경

- 출발점: **[CATMOA](https://github.com/progh2/catmoa)** — 교사 개발자 해커톤
  결과물, 공동 참여
  - 메시지 입력 → 일정 추출 → 캘린더 등록 흐름과 화면 구석 작은 창 형태를 차용
- 재개발 이유: **학교 환경에서 Google API 사용 제한**
  - 학교 서버·네트워크에서 막히면 캘린더 등록 전체가 중단
  - → 일정을 Python 명령으로 직접 다룰 수 있는 **Outlook 버전으로 재개발**
  - Outlook은 로컬 COM 조작이므로 외부 API 차단과 무관
- 변경점
  - **UI 단순화**: 작은 창에서 입력란·실행 버튼 제거하고 상태만 표시,
    검토창은 카드에서 바로 `등록`/`미등록`
  - **[Upstage Solar Pro 4](https://console.upstage.ai/)** 로 분석·테스트
    (교사 무료). API 키 없으면 오프라인 규칙으로 동작
- 원본 CATMOA 저장소에는 변경 가하지 않음

---

## 빠른 시작

### 실행 환경

| 항목 | 요구사항 |
|---|---|
| OS | Windows 10 / 11 |
| Python | 3.11 이상 |
| Outlook | **클래식** 데스크톱 앱, 로그인 상태 (New Outlook 단독 불가) |
| Upstage API 키 | 선택. 없으면 오프라인 규칙 분석만 |

### 설치

```powershell
python -m pip install -r requirements.txt
python app.py
```

| 패키지 | 용도 |
|---|---|
| `pywin32` | Outlook COM, 전역 단축키, 트레이 |
| `openai` | Upstage 호환 클라이언트 |
| `Pillow` | 아이콘 변환, 이미지 |
| `pdfplumber`, `olefile` | PDF·HWP 텍스트 추출 |
| `tkinterdnd2` | 드래그앤드롭 |
| `pytesseract` | 이미지 OCR (선택) |

### API 키

`.env` 또는 환경변수.

```text
UPSTAGE_API=up_...
UPSTAGE_MODEL=solar-pro4
```

- 인식 이름 우선순위: `UPSTAGE_API` → `UPSTAGE_API_KEY` → `SOLAR_API_KEY`
- `.env`는 로컬 전용, Git 제외

### 사용자 프로필

- `담임`/`업무` 판정에 사용자 역할 정보 필요
- `profile.example.json` 복사 → `profile.json` 작성 (Git 제외)

```json
{ "role": "예: 중학교 교사", "department": "", "subject": "", "notes": "" }
```

### 첫 사용

1. 다른 앱에서 업무 메시지 선택
2. **Ctrl+Shift+X** → 작은 창이 `처리 중`으로 전환
3. 분석 완료 시 검토창 자동 열림
4. 카드의 제목·날짜 확인 후 `등록` / `미등록`
5. 남은 카드와 분석이 모두 없으면 창 자동 종료

---

## 화면

### 작은 상태 창

- 크기 `180×128`, 작업 표시줄 바로 위 오른쪽 끝에 밀착
- 입력란·실행 버튼 없음. 현재 상태만 표시
- 실행 중 항상 위
- 제목 표시줄 없음 → 최소화·닫기 버튼 없음
  - 이동: 위쪽 `업무 정리` 줄을 잡고 드래그
  - 숨기기·종료: 트레이 아이콘
  - 트레이 생성 실패 시 제목 표시줄 유지 (창으로 종료 가능)

| 상태 | 뜻 |
|---|---|
| `대기 중` | 입력 대기 |
| `처리 중 n건` | Solar Pro 4 분석 중 |
| `대기 n건` | 분석 완료, 검토창 여는 중 |
| `검토 중` | 검토창 열림 (분석 남으면 `분석 n` 병기) |
| `일시중지` | 새 입력 받지 않음 |

- 창 클릭 → 파일 선택
- 창에 파일 드롭 → 즉시 읽기
- `⋯` 메뉴 → 상세 상태 · Outlook 동기화 상태 · 선택 텍스트 가져오기 ·
  붙여넣기 · 파일 선택 · 일시중지 · Outlook 연결 확인 · 단축키 안내 · 종료

### 트레이 아이콘

- 오른쪽 클릭 → `창 보이기` / `일시중지`·`재개` / `종료`
- 왼쪽 클릭 → 창을 앞으로
- 일시중지 중: 단축키·드롭·붙여넣기 입력을 대기열에 넣지 않음.
  이미 분석 중인 항목은 그대로 완료
- 아이콘 원본은 `tray_icon.png`
  - 첫 실행 때 `.ico`로 변환해 임시 폴더에 캐시
  - 창 제목 표시줄 · 작업 표시줄 · 알림 영역이 동일 그림 사용
  - 원본 교체 시 다음 실행에서 자동 재생성

### 검토창

- 크기 `760×620`, 작은 창을 가리지 않도록 화면 왼쪽 위 배치
- 왼쪽: 원문 (기본 수정 잠금, `원문 수정`으로 해제)
- 오른쪽: 저장할 내용. 메시지별 `메시지 1`, `메시지 2` 구분
- 오른쪽 위 `⋯`: 기준일 설정 · 다시 분석 · 오프라인 규칙으로 분석 ·
  JSON 복사 · JSON 저장
- 아래: `닫기`, `모두 등록 n건`
- 검토 중에도 새 메시지 입력 가능
  - 분석이 끝난 메시지는 대기 없이 같은 창에 추가
  - 헤더에는 분석 중인 건수만 `분석 중 n건` 표시
- 한 메시지의 카드를 모두 처리하면 초점·왼쪽 원문이 다음 메시지로 이동
- 메시지 헤더 클릭 시 해당 메시지 원문으로 전환
- 남은 카드와 분석이 모두 없으면 창 종료

### 카드

카드 1개 = Outlook 항목 1개.

```text
┌──────────────────────────────────────────────┐
│ 제목 (길면 줄바꿈, 즉시 편집)                   │
│ [마감] 2026-09-21   시각 ____      등록 대기   │
│ 시청각실 · 조건부        [상세] [미등록]  [등록] │
└──────────────────────────────────────────────┘
```

| 요소 | 동작 |
|---|---|
| 제목 | 카드 폭 전체 사용, 길면 최대 4줄까지 줄바꿈. 즉시 편집 |
| 날짜/시각 | 상세를 열지 않고 즉시 편집. 수정 시 등록 가능 여부 재계산 |
| `등록` | 그 카드만 즉시 Outlook 기록. 오클릭 방지를 위해 `미등록`과 간격 |
| `미등록` | 카드 접어 둠. 다시 누르면 `되돌리기` |
| `상세` | 분류 · 일시 · 마감 · 마감 시각 · 종일 · 장소 · 알림 · 반복 · 마감일 캘린더 표시 · 내역 |
| `모두 등록` | 분석이 안전하다고 본 항목만 일괄. 조건부·확인 필요 항목은 제외 |

| 상태 라벨 | 뜻 |
|---|---|
| `등록 대기` | 즉시 등록 가능 |
| `조건부 · 선택 시 등록` | 해당자만. 기본 선택 해제 |
| `확인 필요` | 수행 시각 등 누락으로 등록 차단. 수정 시 해제 |
| `대상 확인 필요` | 본인 해당 여부 불명확 |
| `중복 제외 · 사유` | 기존 항목과 겹침. 제목·시간 수정 시 해제 |
| `등록 실패 · 사유` | Outlook 기록 실패. 수정 후 재시도 |
| `등록 완료` | 완료. 버튼 잠금 |

- 참조·완료·중복 항목은 카드로 펼치지 않고 `참조/제외 n건 보기`에 접어 사유 표시

---

## 입력 방법

### 1. 선택 텍스트 (Ctrl+Shift+X)

Windows에 "다른 앱의 현재 선택 텍스트"를 읽는 API가 없어 다음 순서로 동작.

1. 현재 클립보드 내용 보관
2. 눌린 Ctrl/Shift/X 해제 후 Ctrl+C 합성 전송
3. 클립보드가 바뀌면 그 텍스트 사용
4. 원래 클립보드 내용 복원

- 선택이 없으면 현재 클립보드 사용. 둘 다 없으면 작은 창에 안내
- 관리자 권한 프로그램은 Windows UIPI 정책으로 합성 Ctrl+C가 전달되지 않을 수 있음
  → 직접 Ctrl+C 복사 후 단축키

### 2. 붙여넣기 (Ctrl+V)

- 작은 창 클릭으로 포커스 → Ctrl+V
- `⋯ → 붙여넣기` 동일 동작

### 3. 파일

- 작은 창에 드롭, 또는 창 클릭으로 파일 선택

| 종류 | 형식 |
|---|---|
| 문서 | PDF, HWP, HWPX, DOCX |
| 텍스트 | TXT, Markdown, CSV, TSV, LOG |
| 이미지 | PNG, JPG, JPEG, BMP, GIF, WEBP, TIFF |

- PDF/HWP 계열: 로컬 텍스트 추출 후 분석
- 이미지: `pytesseract` + Windows Tesseract 실행 파일 + 한국어 언어팩 설치 시에만
  읽음. 미설치 시 텍스트 붙여넣기 안내

### 4. `key=value`

```text
title=학부모 상담
target=calendar
start=2026-09-10 14:00
duration_minutes=30
location=상담실
reminder_minutes=15
memo=상담 자료 준비
```

- `target=both` → 같은 입력에서 일정과 작업을 함께 검토

---

## 분석

### 입력 예시

```text
2026-09-10 14:00 학부모 상담 장소: 상담실 알림 15분 전
2026-09-12 운동회 행사 장소: 운동장
2026-09-15 학생 명단 제출
```

- 시간 있는 회의·상담 → Outlook 일정
- 제출·확인·작성 → Outlook 작업
- 날짜만 있는 Calendar 후보 → 00:00/종일로 임의 변환하지 않고 `확인 필요` 유지

### 행동 단위 묶기

모집·희망자·신청 공지는 날짜별로 나누지 않고 **실제 사용자 행동 단위**로 묶음.

| 입력 | 결과 |
|---|---|
| `토익 시험 감독관 모집` + `시험일 2026-08-23 08:30` + `당일 오전 10시까지 접수` | `토익 시험 감독관 신청` Todo **1건** |

- 신청은 조건부 → 기본 선택 해제
- 시험일 → 관련 행사 정보(`event_context`)로 보존
- 날짜 미확정 10시 → 마감 시각 문맥(`due_time`)으로 보존
- 시험 참석 일정, `접수 마감 확인` 작업은 자동 생성 안 함

### 1차 분류 (scope) — LLM 판단

| 분류 | 기준 | Outlook |
|---|---|---|
| 담임 | 최종 대상이 **학생**. 전달·안내·배부·지도 | 등록 |
| 업무 | 최종 대상이 **본인**. 제출·참석·이수·신청 | 등록 |
| 참조 | 읽고 넘기면 되는 안내 | 등록 안 함 (SQLite만) |

### 2차 분류 (action) — 코드가 필드에서 파생

| 조건 | action | Outlook |
|---|---|---|
| 마감(`due`) 있음 | `todo_with_due_date` | 작업 + 마감일 + 내역, 마감일 캘린더 표시 |
| 일시(`at`) 있음 | `todo_with_date` | 캘린더 + 알림 |
| 둘 다 없음 | `todo` | 작업 (마감 없음) |
| scope=참조 | `none` | 등록 안 함 |

- LLM에 2차 라벨까지 요청 시 실측 **7%가 자기모순**
  (`todo_with_due_date` 인데 `due=null`)
- 사실만 받아 코드에서 파생 → 해당 모순이 구조적으로 불가능
- 검토창에서 분류·일시·마감 수정 시 라벨 즉시 재파생

### 반복·의미 필드

- 반복: `repeat_freq`(daily/weekly/monthly) + `repeat_detail`("화,목")
  → Outlook 반복 작업·일정으로 등록
- 보존 필드: `intent`(행동), `applicability`(조건부 여부),
  `source_state`(완료/안내), `condition`, `due_time`, `temporal_context`,
  `checklist`, `evidence`, `review_issues`
- 날짜가 여러 개여도 `event_context`·`external_deadline`·`historical` 역할로
  분리 → 주 행동의 등록 대상과 섞이지 않음

### Solar Pro 4 호출

| 항목 | 값 |
|---|---|
| endpoint | `https://api.upstage.ai/v1` (OpenAI 호환) |
| 모델 | `solar-pro4` |
| 타임아웃 | 첫 요청 60초, 재시도 90초 |
| 최대 시도 | 2회 |
| 추론 | `reasoning_effort="none"` |
| 프롬프트 | `semantic-v4-compact` (API 필수 필드 22개 → 12개) |

- 전부 실패 → 오프라인 규칙 분석으로 전환, 검토창에 표시
- 간소화 JSON Schema + 로컬 형식 검증 병행
- 빈 선택 필드·내부 ID·기본값은 앱에서 복원
- 분석은 백그라운드 FIFO 대기열 → 새 메시지 계속 입력 가능
- 오프라인 전환 결과는 기본 선택 해제 → 등록 전 조건·확인 사유 확인 필요

### 기준일

- `내일`·`다음주 월요일` 계산 기준. 기본값 오늘
- 검토창 `⋯ → 기준일 설정`에서 변경
- 메시지 발송일 입력 시 정확도 상승
- 본문에 발송일·작성일이 있으면 모델이 우선하도록 프롬프트에 규칙 포함

### 등록 전 차단 항목

- 구조 불일치 JSON, 필수 필드 누락 응답
- **원문에 없는 수행 시각** (모델이 생성한 `00:00` 등)
- 시각 미정 Calendar 후보 (`start=null`)
- 조건 없는 조건부 항목
- 대상 불명확 항목 (`applicability=unknown`)
- Outlook 중복 조회 실패 → 등록 보류

---

## Outlook 등록

### 등록 대상

| 분석 결과 | Outlook |
|---|---|
| 일시 있는 업무 | Calendar 약속 (+ 알림) |
| 마감일 있는 업무 | 작업(Task) + 마감일, **+ 마감일 당일 종일 표시** |
| 날짜 없는 할 일 | 작업(Task) |
| 참조 | 생성 안 함. SQLite만 |

- 시간 일정에 종료가 없으면 저장 시 60분 적용
- Todo는 `due=null`이어도 등록 가능. `due_time`은 설명에 보존

### 마감일 캘린더 표시

Outlook 작업(Task)은 캘린더에 표시되지 않음 → 마감 있는 작업은 마감일 당일에
종일 표시를 추가 생성.

- 제목 `[마감] 원래 제목`, 범주 `마감`
- 하루를 차지하지만 **`한가함(Free)`이며 중복·시간 겹침 검사에서 제외**
  - 제외하지 않으면 마감일 하루가 통째로 막혀 그날 실제 일정 등록이 전부 차단됨
- 항목 1개 → Outlook 항목 2개(작업 + 표시)
  - 불필요 시 카드 `상세`의 `마감일을 캘린더에도 표시` 해제
- 표시 생성 실패해도 이미 만들어진 작업은 유지, 사유만 알림
- 같은 제목·같은 날짜 표시가 있으면 재생성 안 함
- 일시 있는 일정, 마감 없는 할 일에는 생성 안 함

### 중복·시간 겹침

카드 표시 전과 등록 직전 2회 확인.

**Calendar**

- 기존 항목과 시간 구간 겹치면 제외 (종료 없으면 60분으로 간주)
- 종일 일정은 하루 단위 비교
- 같은 메시지에서 만든 일정끼리 겹치면 뒤의 것 제외

**Todo**

- 제목의 공백·구두점 정리 후 비교
- 제목 + 마감일 일치 시 제외
- 마감일 없는 동일 제목 Todo도 제외

제외 카드는 사유 표시. 제목·시간 수정 시 제외 해제되고 재등록 가능.

### Outlook 로컬 미러

카드마다 등록하는 화면에서는 클릭할 때마다 폴더 전체를 COM으로 훑을 수 없음
(항목별 다중 속성 조회 → 일정·작업이 많은 프로필에서 지연 발생).
→ 현재 프로필을 SQLite에 복제하고 중복 검사는 사본만 조회.

| 시점 | 동작 |
|---|---|
| 앱 시작 | 전용 COM 스레드가 Calendar/작업 전체 1회 읽어 미러 생성 |
| 5분마다 | 동일 스레드가 재조회, 추가·변경·삭제 반영 |
| 이 앱이 등록할 때 | 생성 항목을 미러에도 즉시 기록 (write-through) |

- 미러가 오래되면 다음 검사에서 백그라운드 갱신 요청, 결과는 대기하지 않음
- 첫 동기화 완료 전에는 Outlook 직접 조회
- **한계**: 다른 프로그램이 만든 항목은 다음 동기화까지 미러에 없음
  → 최대 5분의 사각. 클릭 반응성과의 트레이드오프
- 동기화 실패 시 `⋯` 메뉴에 표시
- 읽기 전용. Outlook 항목 변경 없음. 미러는 로컬 SQLite에만 존재

### 연결 전제

- `pywin32`의 `Outlook.Application` COM 사용 → **클래식 Outlook 설치 필요**
- New Outlook 단독 환경에서는 동작 안 함
- Outlook 로그인 상태로 실행
- `⋯ → Outlook 연결 확인`은 연결만 검사, 항목 생성 안 함
- 로컬 COM 자동화이므로 Microsoft Graph 등 외부 HTTP API 미사용
- 동일 Microsoft 계정의 Outlook 작업은 Microsoft To Do와 동기화될 수 있음

---

## 저장 데이터

### SQLite

파일: `data/work_items.sqlite3`

| 테이블 | 내용 |
|---|---|
| `inputs` | 원문, 생성 시간, 기준일·발송일·시간대 등 `context_json` |
| `work_items` | 분류, 제목, 시간/마감일, 장소/범주/알림, 설명, `saved`, `outlook_entry_id`, `calendar_entry_id`, 오류, `due_time`, `semantic_json` |
| `outlook_mirror` | Outlook 프로필의 Calendar/작업 사본 (중복 검사용) |
| `sync_state` | 마지막 미러 동기화 시각 |

```text
pending → saved
pending → failed
saved   → deleted   # 별도 관리 파이프라인에서 삭제된 경우
```

- 분석 직후 pending 기록 → Outlook 성공 시 `saved=1` + `outlook_entry_id`,
  실패 시 `last_error`
- 등록 성공 카드는 버튼 잠금 → 이중 등록 방지
- 스키마 변경 시 앱 시작 때 컬럼 자동 추가. 기존 기록 유지

### JSON 내보내기

검토창 `⋯ → JSON 복사` / `JSON 저장` 출력 구조.

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

---

## 검증 결과

### 로컬 테스트 (API 호출 없음)

```powershell
python selftest.py --offline                # 42/42
python -m unittest test_review_flow         # 12/12
python -m unittest test_compact_extraction  # 7/7
python -m compileall -q .
python -m pip check
```

- `selftest.py`는 Outlook에 항목을 만들지 않음
- 단, 전역 단축키 등록을 검사하므로 **앱 실행 중에는 해당 항목 1건 실패**
  → 앱 종료 후 실행

### API 사용 검사 (비용 발생)

```powershell
python selftest.py                          # 30/30 (미러 추가 전 측정)
python benchmark_latency.py --repeats 2 --baseline eb0f898
```

### 의미 정확도

회귀 fixture `test_set_30.xlsx`의 `테스트셋` 시트 본문 32건 기준.
신규 메시지 전체의 일반화 점수가 아니라 프롬프트·의미 필드의 기준선.

> **아래 수치는 프롬프트 간소화 이전 버전 결과이며 현재 버전의 정확도가 아님.**
> 축약 이후 속도만 재측정, 의미 정확도는 미재채점.

| 항목 | 값 |
|---|---|
| 응답 확보 | 31/32건 (1건은 2회 시간 초과 후 fallback) |
| 응답시간 평균 / 중앙값 | 31.92초 / 27.62초 |
| 95백분위 / 최대 | 92.03초 / 97.53초 |
| 60초 이내 / 90초 이내 | 29건 / 29건 |
| 원문에 없는 `00:00` Calendar | 0건 |
| 수동 감사 | 완전 12건 / 부분 11건 / 개선 필요 9건 |

- 프롬프트 축소 후 합성 메시지 8회: 평균 12.178초 → 4.606초
- 완료 안내 2회 제외한 실제 추출 6회 평균 5.595초 → **목표 5초 미달**
- 상세: [LATENCY_REPORT.md](LATENCY_REPORT.md)

반복 실행 편차

| 사례 | 결과 |
|---|---|
| 토익 모집 | 3회 모두 `신청 Todo 1건` 안정 |
| 충전·책상 배열 | 3회 중 2회만 2건 분리 유지 |
| 연수·감독 | 조건부 처리, 긴 공지의 행동 단위 분리가 회차마다 변동 |

### 종합 판정

- 신규 등록 UI·DB·대기열: **구현 완료**
- 의미 추출: **검토형 베타**
- 후보가 많거나 조건이 복잡한 메시지는 카드의 원문·확인 사유를 함께 확인 후 등록
- 무검토 자동 등록은 비권장

---

## 개발자용

### 파일 구성

**실행 흐름**

| 파일 | 역할 |
|---|---|
| `app.py` | 작은 상태 창, 전역 단축키, 파일 드롭, FIFO 분석 대기열, 트레이 연결 |
| `review.py` | 검토창. 왼쪽 원문 / 오른쪽 메시지별 카드 목록, JSON 내보내기 |
| `cards.py` | WorkItem 카드. 즉시 편집 + 카드 단위 등록/미등록 |
| `service.py` | 분류 → 중복 검사 → SQLite pending → Outlook 기록 → saved/failed |
| `models.py` | `WorkItem`, scope/action 파생 규칙, 등록 가능 판정, 검증 |
| `storage.py` | SQLite 저장소, 스키마 마이그레이션 |

**Outlook**

| 파일 | 역할 |
|---|---|
| `outlook_adapter.py` | 클래식 Outlook COM 어댑터. 일정·작업·마감 표시 생성, 내역 조회 |
| `outlook_mirror.py` | 프로필 로컬 사본, 백그라운드 동기화 스레드, write-through |

**분석**

| 파일 | 역할 |
|---|---|
| `upstage_classifier.py` | Solar Pro 4 구조화 JSON 분류, 응답 검증 |
| `compact_extraction.py` | 간소화 프롬프트·API 스키마·축약 응답 복원 |
| `classifier.py` | 오프라인 규칙 분류 |

**입력·플랫폼**

| 파일 | 역할 |
|---|---|
| `selection.py` | 선택 텍스트 캡처, 클립보드 읽기/쓰기/복원 |
| `hotkey.py` | `RegisterHotKey` 전역 단축키 (전용 스레드 + 메시지 루프) |
| `tray.py` | 알림 영역 아이콘, 일시중지/종료 메뉴 (전용 스레드 + 메시지 루프) |
| `windows_drop.py` | Windows 파일 드래그앤드롭 브리지 |
| `file_input.py` | 문서·이미지 로컬 텍스트 추출 |
| `icons.py` | `tray_icon.png` → `.ico` 변환, 창·작업 표시줄 아이콘 |
| `theme.py` | 공통 색·글꼴·위젯 스타일 |

**검증**

| 파일 | 역할 |
|---|---|
| `selftest.py` | Outlook에 쓰지 않는 로컬·연결 검증 |
| `test_review_flow.py` | Tk 위젯 조작 기반 검토창 회귀 검사 |
| `test_compact_extraction.py` | 축약 응답·조건·반복·시간 차단 회귀 검사 |
| `benchmark_latency.py` | 합성 메시지 API 속도·핵심 의미 비교 (API 호출) |

- 의존 방향 단방향: UI → service → 분류/저장/COM → models
- `models.py` 수정 → 거의 전 모듈 영향
- `cards.py` / `review.py` 수정 → UI 내부에서 종결

### 커스터마이징

| 대상 | 위치 |
|---|---|
| 단축키 | `app.py`의 `VK_X`, `HOTKEY_LABEL`, `HotkeyListener(modifiers=...)` |
| 아이콘 | `tray_icon.png` 교체 (투명 배경·정사각형 권장. 옅은 알파 배경은 자동 제거) |
| 미러 동기화 주기 | `outlook_mirror.py`의 `SYNC_INTERVAL_SECONDS`(기본 300초), `STALE_AFTER_SECONDS` |
| 창 위치·크기 | `app.py`의 `WINDOW_WIDTH`, `WINDOW_HEIGHT`, `EDGE_MARGIN` |

---

## 문제 해결

| 증상 | 원인 · 조치 |
|---|---|
| 단축키 등록 실패 표시 | 다른 프로그램이 Ctrl+Shift+X 사용 중. 해당 프로그램 종료 또는 `app.py`에서 조합 변경. 그동안 `⋯` 메뉴의 붙여넣기·파일 선택은 사용 가능 |
| Ctrl+Shift+X 무반응 | 상대 프로그램이 관리자 권한(UIPI). 직접 Ctrl+C 복사 후 재시도 |
| `Outlook 동기화 실패` | 클래식 Outlook 실행·로그인 상태 확인. 실패해도 첫 등록 시 Outlook 직접 조회 |
| `중복 확인 실패로 보류` | Outlook 조회 실패. 중복 검사는 안전 전제라 생략하지 않음. 연결 확인 필요 |
| 이미지 인식 안 됨 | Tesseract 실행 파일 + 한국어 언어팩 필요. 미설치 시 텍스트 붙여넣기 |
| 트레이 아이콘 없음 | 작은 창이 제목 표시줄 유지 → `X`로 종료 가능. `⋯` 메뉴에도 일시중지·종료 있음 |
| 작은 창 이동 | 위쪽 `업무 정리` 줄 드래그 (제목 표시줄이 없어 그 줄이 손잡이) |
| 마감일마다 캘린더가 막힘 | 마감 표시는 `한가함` + 겹침 검사 제외이므로 해당 없음. 다른 프로그램이 만든 종일 일정인지 확인 |
| 작업 표시줄에 파이썬 아이콘 | `python app.py`로 실행했는지 확인. `main()`에서 앱 식별자 설정 |

---

## 개인정보

- 원문은 SQLite `inputs.original_text`에 **원본 그대로** 저장
- 현재 요구사항에 따라 **PII 탐지·마스킹 미수행**
- API 키가 있으면 원문이 Upstage API로 전송. 키가 없으면 외부 전송 없음
- `.env`, `profile.json`, `data/*.sqlite3`, `test_set_30.xlsx`는 `.gitignore` 제외 대상
  → 공유 금지

---

## 범위 밖

- 기존 Outlook 항목 조회·수정·삭제 UI (후속 작업 보류)
- 쿨메신저 UDB 자동 탐색
- 개인정보 탐지·마스킹
- New Outlook / Microsoft Graph 지원
- 메일 발송

진행 상태와 다음 작업 후보: [WORK_TREE.md](WORK_TREE.md)
