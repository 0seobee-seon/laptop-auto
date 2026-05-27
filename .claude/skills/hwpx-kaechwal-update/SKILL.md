---
name: hwpx-kaechwal-update
description: |
  HWPX 주간업무 보고서에서 오늘 날짜 기준으로 개찰일이 도래한 '진행중' 항목을
  '개찰' 섹션으로 자동 이동하고 업데이트된 파일을 저장합니다.

  다음 상황에서 반드시 이 스킬을 사용하세요:
  - "진행중 항목 개찰로 올려줘"
  - "오늘 기준으로 개찰일 된 것 업데이트해줘"
  - "주간업무 hwpx 개찰 현황 최신화해줘"
  - "practice.hwpx 업데이트해줘"
  - HWPX 파일과 함께 개찰·진행중 섹션 변경 요청이 오는 경우
  - 날짜를 명시하든 않든, 주간업무 표 상태 변경 요청이면 이 스킬로 처리하세요.
---

# HWPX 개찰 현황 업데이트

## 목적
주간업무 보고서(.hwpx) 내 표에서 **오늘 날짜(또는 지정 날짜) 이하**의 개찰일을 가진
'진행중' 항목을 '개찰' 섹션으로 이동하고, 행 병합(rowSpan)을 자동 재계산합니다.

## 전제 조건
- Python 3.8+ (표준 라이브러리만 사용, 추가 설치 없음)
- 입력 파일: `.hwpx` 형식
- 표 구조: 구분(col0), 번호(col1), 용역명(col2), 단장(col3), 제출일(col4),
  발표/면접(col5), **개찰일(col6)**, 용역비(col7), 내용(col8)

## 입출력
| 항목 | 값 |
|------|-----|
| 입력 | 사용자가 지정한 `.hwpx` (기본: `practice.hwpx`) |
| 출력 | `<원본명>_update.hwpx` (기본: `practice_update.hwpx`) |
| 기준 날짜 | 오늘 날짜 자동 적용, `--date YYYY-MM-DD`로 지정 가능 |

---

## 워크플로우

### Step 1 — 스크립트 실행

```bash
python .claude/skills/hwpx-kaechwal-update/scripts/update_kaechwal.py \
  --input  practice.hwpx \
  --output practice_update.hwpx
```

날짜를 직접 지정하려면:
```bash
python .claude/skills/hwpx-kaechwal-update/scripts/update_kaechwal.py \
  --input  practice.hwpx \
  --output practice_update.hwpx \
  --date   2026-05-21
```

### Step 2 — 스크립트 결과 확인

스크립트는 다음 JSON을 stdout으로 출력합니다:
```json
{
  "success": true,
  "reference_date": "2026-05-21",
  "moved_items": [
    {"num": "4", "name": "의정부법조타운 S-2BL", "kaechwal_date": "5/21"}
  ],
  "remaining_items": [
    {"num": "7", "name": "쌍문1동 공공복합청사", "kaechwal_date": "5/27"}
  ],
  "kaechwal_count": 7,
  "jinhang_count": 1,
  "output_path": "practice_update.hwpx"
}
```

`success: false`이면 `error` 필드를 확인하고 사용자에게 보고합니다.

### Step 3 — 결과 검증 (kordoc)

```bash
npx --yes --package kordoc --package pdfjs-dist kordoc practice_update.hwpx --format json
```

kordoc 결과에서 표 구조를 확인합니다:
- 개찰 섹션의 rowspan = 이동된 항목 수 + 기존 개찰 수 + 1(빈 행)
- 진행중 섹션의 rowspan = 유지된 항목 수 + 1(빈 행)
- 개찰일 ≤ 기준날짜인 항목이 모두 개찰 섹션에 있음
- '추후' 항목은 진행중에 유지됨

### Step 4 — 사용자에게 결과 보고

이동된 항목과 유지된 항목을 표로 정리해 보고합니다.

---

## 날짜 파싱 규칙

| 형식 | 예시 | 처리 |
|------|------|------|
| M/D | `5/21` | 당해 연도 날짜로 파싱 |
| `추후` | 추후 | None → 이동 안 함 (진행중 유지) |
| 빈 값 / `-` | | None → 이동 안 함 |

---

## 오류 처리

| 오류 상황 | 대응 |
|-----------|------|
| 진행중 섹션 없음 | `error` 필드 확인 후 사용자 안내 |
| 이동할 항목 없음 | `message` 필드로 안내, 파일은 생성 안 함 |
| HWPX 파싱 실패 | 스택 트레이스 포함하여 사용자에게 보고 |
| 출력 경로 쓰기 불가 | 경로 권한 확인 후 재시도 |

---

## 검증 서브에이전트 호출 조건

다음 경우에 `.claude/agents/validator/AGENT.md`의 검증 에이전트를 호출합니다:
- 스크립트 실행 성공 후 산출물 품질 확인이 필요할 때
- 사용자가 "검증해줘" / "확인해줘"를 명시적으로 요청할 때
