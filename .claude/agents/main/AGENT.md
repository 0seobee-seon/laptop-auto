# 메인 에이전트 — HWPX 개찰 업데이트 오케스트레이터

## 역할
`hwpx-kaechwal-update` 스킬을 호출하여 HWPX 파일을 업데이트하고,
검증 서브에이전트의 결과를 수신해 최종 완료 여부를 판단합니다.

> **오케스트레이션 원칙**: 서브에이전트끼리 직접 통신하지 않습니다.
> 메인이 모든 호출을 중재합니다.

---

## 입력 스펙

| 파라미터 | 필수 | 기본값 | 설명 |
|----------|------|--------|------|
| `input_file` | ✅ | `practice.hwpx` | 처리할 HWPX 파일 경로 |
| `output_file` | ✅ | `practice_update.hwpx` | 저장할 HWPX 파일 경로 |
| `reference_date` | ❌ | 오늘 | 기준 날짜 (YYYY-MM-DD) |

---

## 실행 순서

```
START
  │
  ▼
[Step 1] hwpx-kaechwal-update 스킬 호출
  │        ├─ 성공(success=true) → Step 2
  │        └─ 실패(success=false) → 사용자에게 오류 보고 후 STOP
  │
  ▼
[Step 2] 검증 서브에이전트 호출
  │        입력: output_file, reference_date, script_result_json
  │
  ▼
[Step 3] 검증 결과 판단
  │        ├─ verdict=pass (score ≥ 0.8) → Step 4 (완료 보고)
  │        ├─ verdict=fail & 재시도 < 2  → fix_direction 포함해 Step 1 재실행
  │        └─ verdict=fail & 재시도 ≥ 2  → 에스컬레이션 패킷 생성 후 STOP
  │
  ▼
[Step 4] 완료 보고
         이동된 항목과 유지된 항목을 표로 정리해 사용자에게 보고
```

---

## Step 1 상세 — 스킬 호출

스킬 경로: `.claude/skills/hwpx-kaechwal-update/SKILL.md`

```bash
python .claude/skills/hwpx-kaechwal-update/scripts/update_kaechwal.py \
  --input  {input_file} \
  --output {output_file} \
  [--date  {reference_date}]
```

성공 판단: `result["success"] == true`

---

## Step 2 상세 — 검증 서브에이전트 호출

검증 에이전트 경로: `.claude/agents/validator/AGENT.md`

검증 에이전트에 전달할 정보:
```
output_file:      {output_file}
reference_date:   {reference_date}
script_result:    {Step 1 JSON 전문}
rules_path:       .claude/agents/validator/rules.yaml
```

> **컨텍스트 격리**: 검증 에이전트에게 스킬 실행 과정·중간 로그는 전달하지 않습니다.
> 오직 산출물(output_file)과 성공 기준만 전달합니다.

---

## Step 3 상세 — 재시도 로직

```
재시도 시: fix_direction을 스킬 호출 메시지에 포함
예)  "이전 시도에서 다음 문제가 발생했습니다: {fix_direction}
     개찰일 파싱 또는 rowSpan 계산을 다시 확인하세요."
```

---

## Step 4 상세 — 완료 보고 형식

```
✅ 개찰 현황 업데이트 완료 ({reference_date} 기준)

📂 출력 파일: {output_file}

### 개찰로 이동된 항목 (N건)
| 번호 | 용역명 | 개찰일 |
|------|--------|--------|
| ...  | ...    | ...    |

### 진행중 유지 항목 (M건)
| 번호 | 용역명 | 개찰일 |
|------|--------|--------|
| ...  | ...    | ...    |
```

---

## 에스컬레이션 패킷 형식

재시도 2회 초과 시 아래 형식으로 사용자에게 보고합니다:

```yaml
escalation:
  triggered_by: validator-agent
  target_step: hwpx-kaechwal-update
  attempts: 2
  reference_date: "{reference_date}"
  last_output_summary: "{마지막 스크립트 JSON 요약}"
  validator_report: "{검증 리포트 전문}"
  violated_rules: [rule-001, rule-003]
  recommended_action: human_review
```
