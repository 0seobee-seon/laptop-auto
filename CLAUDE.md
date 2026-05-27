# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## 언어
모든 응답은 **한국어**로 작성한다.

---

## Behavioral Guidelines

Reduce common LLM coding mistakes. Merge with project-specific instructions as needed.
Tradeoff: These guidelines bias toward caution over speed. For trivial tasks, use judgment.

### 1. Think Before Coding

Don't assume. Don't hide confusion. Surface tradeoffs.

Before implementing:
- State your assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them — don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.
- If something is unclear, stop. Name what's confusing. Ask.

### 2. Simplicity First

Minimum code that solves the problem. Nothing speculative.

- No features beyond what was asked.
- No abstractions for single-use code.
- No "flexibility" or "configurability" that wasn't requested.
- No error handling for impossible scenarios.
- If you write 200 lines and it could be 50, rewrite it.

Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

### 3. Surgical Changes

Touch only what you must. Clean up only your own mess.

When editing existing code:
- Don't "improve" adjacent code, comments, or formatting.
- Don't refactor things that aren't broken.
- Match existing style, even if you'd do it differently.
- If you notice unrelated dead code, mention it — don't delete it.

When your changes create orphans:
- Remove imports/variables/functions that YOUR changes made unused.
- Don't remove pre-existing dead code unless asked.

The test: Every changed line should trace directly to the user's request.

### 4. Goal-Driven Execution

Define success criteria. Loop until verified.

Transform tasks into verifiable goals:
- "Add validation" → "Write tests for invalid inputs, then make them pass"
- "Fix the bug" → "Write a test that reproduces it, then make it pass"
- "Refactor X" → "Ensure tests pass before and after"

For multi-step tasks, state a brief plan:
```
1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
```

Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

---

## 프로젝트 개요

미래사업팀 주간업무 보고서(`.hwpx`)의 수행 Project 표를 관리하는 자동화 워크플로우.
오늘 날짜 기준으로 개찰일이 도래한 '진행중' 항목을 '개찰' 섹션으로 이동하고 파일을 저장한다.

---

## 파일 구조

```
/project-root
  ├── CLAUDE.md                          ← 이 파일
  ├── practice.hwpx                      ← 원본 입력 파일
  ├── practice_update.hwpx               ← 자동화 산출물
  │
  └── .claude/
      ├── skills/
      │   └── hwpx-kaechwal-update/      ← 재사용 가능한 스킬
      │       ├── SKILL.md               ← 스킬 정의 (트리거·워크플로우)
      │       └── scripts/
      │           └── update_kaechwal.py ← 핵심 변환 스크립트 (Python)
      │
      └── agents/
          ├── main/
          │   └── AGENT.md               ← 오케스트레이터 에이전트
          └── validator/
              ├── AGENT.md               ← 검증 전담 에이전트
              └── rules.yaml             ← 누적 검증 규칙 (버그 이력 기반)
```

---

## 핵심 스킬

### `hwpx-kaechwal-update`

| 항목 | 내용 |
|------|------|
| 트리거 | "개찰로 올려줘", "진행중 항목 업데이트", "주간업무 최신화" 등 |
| 입력 | `.hwpx` 파일 |
| 출력 | `*_update.hwpx` (행 재배치·rowSpan 재계산 완료) |
| 스크립트 | `update_kaechwal.py` — Python 표준 라이브러리만 사용 |

**스크립트 직접 실행:**
```bash
python .claude/skills/hwpx-kaechwal-update/scripts/update_kaechwal.py \
  --input  practice.hwpx \
  --output practice_update.hwpx
# 날짜 지정 시: --date 2026-05-21
```

---

## 에이전트 구성

```
사용자 요청
    │
    ▼
[메인 에이전트]  .claude/agents/main/AGENT.md
    │   스킬 호출 → update_kaechwal.py 실행
    │   결과 JSON 수신
    │
    ▼
[검증 서브에이전트]  .claude/agents/validator/AGENT.md
    │   산출물 + 성공 기준만 수신 (컨텍스트 격리)
    │   kordoc으로 출력 파일 파싱 및 SC 평가
    │   rules.yaml 규칙 적용
    │
    ▼
    ├── pass (score ≥ 0.8) → 완료 보고
    ├── fail + 재시도 < 2  → fix_direction 포함 재실행
    └── fail + 재시도 ≥ 2  → 에스컬레이션
```

---

## 워크플로우 단계별 처리 주체

| 단계 | 처리 주체 | 성공 기준 | 검증 방법 |
|------|-----------|-----------|-----------|
| HWPX 파싱 | 스크립트 (Python zipfile) | `success: true` | JSON 반환값 확인 |
| 날짜 비교·분류 | 스크립트 | 모든 항목 올바르게 분류 | 검증 에이전트 SC-002, SC-003 |
| XML 재구성 | 스크립트 | rowSpan·rowAddr 정합성 | 검증 에이전트 SC-005, rule-003 |
| 결과 검증 | 검증 서브에이전트 | score ≥ 0.8 | kordoc + rules.yaml |
| 결과 보고 | LLM (메인) | 표 형식·한국어 | 사용자 확인 |

---

## 도구 의존성

| 도구 | 용도 | 설치 |
|------|------|------|
| Python 3.8+ | 핵심 스크립트 실행 | 사전 설치 필요 |
| kordoc + pdfjs-dist | HWPX → Markdown/JSON 변환·검증 | `npx --yes --package kordoc --package pdfjs-dist kordoc` |
| k-skill-rhwp | 단순 텍스트 편집 (선택) | `npx --yes k-skill-rhwp` |

---

## HWPX 표 구조 (참고)

| colAddr | 열 이름 |
|---------|---------|
| 0 | 구분 (개찰/진행중 레이블, rowSpan) |
| 1 | 번호 |
| 2 | 용역명 |
| 3 | 단장 |
| 4 | 제출일 |
| 5 | 발표/면접 |
| **6** | **개찰일** ← 분류 기준 |
| 7 | 용역비(억원) |
| 8 | 내 용 |

---

## 검증 규칙 관리

버그 발생·수정 시 `.claude/agents/validator/rules.yaml`에 rule을 추가한다.

```yaml
- id: rule-NNN
  added_at: "YYYY-MM-DD"
  origin: bug-fix
  bug_summary: "버그 한 줄 요약"
  criterion: "조건 A이면 결과 B여야 한다"
  check_type: logic   # schema | logic | llm
  severity: critical  # critical | major | minor
```

- `critical` 규칙이 하나라도 fail이면 전체 verdict = fail
- 규칙 삭제는 사람 승인 필수 (주석 처리 후 PR)
- 30회 연속 pass된 규칙은 정기 리뷰 대상
