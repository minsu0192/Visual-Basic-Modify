# VB 매크로 인터랙티브 에디터 — 설계 문서

## 한 줄 요약

AI가 생성한 VBA 매크로를 비개발자가 받아, **코드를 직접 건드리지 않고
파라미터(시트명·행 범위·임계값 등)만 위젯으로 안전하게 바꾸어 다시
복사해 갈 수 있게 하는** 정적 웹 페이지.

## 시나리오 (사용자 여정)

1. 사용자가 ChatGPT/Claude에게 "엑셀 매크로 만들어줘"를 요청해 받은
   VBA 코드를 가지고 있다.
2. 본 도구의 좌측 패널에 코드를 붙여넣는다.
3. 우측 패널에 **사용자가 만져도 되는 값들만** 한국어 라벨이 붙은
   위젯으로 자동 표시된다.
   - 예: `For i = 2 To 100` 의 `100` → `[마지막 행 번호: 100]` 슬라이더
   - 예: `Sheets("매출")` 의 `"매출"` → `[시트 이름: 매출]` 텍스트 입력
   - 예: `If 금액 > 50000 Then` 의 `50000` → `[금액 임계값: 50000]`
4. 사용자가 위젯을 변경하면 좌측 코드가 실시간으로 갱신된다.
5. 사용자는 좌측 코드 전체 복사 → Excel VBA 편집기에 붙여넣어 실행.

## 비목표 (이 도구가 하지 *않는* 것)

- VBA를 실행하거나 검증하지 않음. 결과 코드의 동작은 보장하지 않음.
- 풀 AST를 만들지 않음. 형태(주석/공백/들여쓰기)는 원본 그대로 보존.
- 흐름 제어(`Sub`/`Function`/`If`/`For` 자체의 구조)는 편집 대상이
  아님. 비개발자가 바꿀 수 있는 것은 **리터럴 값**과 **사용자
  주석으로 표시된 파라미터**뿐.
- 동적으로 조립되는 값(`Range("A" & i)`)은 시각화하지 않음.
- LLM 호출 없음 (MVP). 모든 라벨은 주석 + 룰 기반 휴리스틱으로 생성.
- 다중 파일/모듈 관리 없음. 한 번에 한 덩어리의 코드.

## 핵심 디자인 결정

### D1. 풀 파서 대신 라인 토크나이저

VBA를 정확히 파싱하려면 ANTLR 급 문법이 필요하지만 비목표.
대신 **라인 단위 분류 + 라인 안에서 토큰 스캔**.
인식 못 하는 라인은 그대로 통과시키되 위젯 추출만 건너뜀.

### D2. AST 재출력이 아닌 오프셋 기반 텍스트 치환

원본 문자열 `source`는 불변. 추출된 각 편집 토큰은
`[startOffset, endOffset]`을 가지며, 사용자의 변경은
`edits: Map<id, newRaw>`에 누적된다.
렌더링 = 모든 edits를 오프셋 순서로 적용해 새 문자열 생성.

이렇게 하면:
- 주석·공백·들여쓰기·한글 100% 보존
- round-trip 버그 없음
- 되돌리기 = `edits` 항목 제거

### D3. 추출 휴리스틱 + 작성자 어노테이션

기본은 휴리스틱(아래 §추출 규칙)으로 자동 노출.
추가로 AI 프롬프트 측에서 다음 같은 어노테이션을 권장:

```vb
'@param 마지막행 min=1 max=10000
Const LAST_ROW = 100
```

위와 같이 `'@param <라벨> [min=.. max=.. step=.. options=A|B|C]`이
바로 위 줄에 있으면 휴리스틱을 덮어쓴다.
**MVP에서는 어노테이션 없이도 동작하지만, 어노테이션이 있으면 UX가
크게 좋아진다.**

### D4. 안전성 우선 — 의심스러우면 숨긴다

비개발자 대상이므로 **잘못된 위젯이 노출되어 코드를 망가뜨리는 것**보다
**위젯이 적게 노출되는 것**이 안전. 숫자 0, 1, 비교 좌변, Mod/배열
인덱스 같은 구조적 리터럴은 추출 제외 기본.

## 데이터 모델

```ts
type Editable = {
  id: string                 // "lit_42"
  kind: "number" | "string" | "bool" | "range" | "forBound" | "ifRhs"
  range: [number, number]    // source 내 오프셋 [start, end)
  raw: string                // 원본 토큰 텍스트 ("100", "\"매출\"")
  value: number | string | boolean
  context: {
    sub: string | null              // 소속 Sub/Function 이름
    line: string                    // 해당 라인 텍스트
    leadingComment: string | null   // 바로 위 ' 주석 (라벨 후보)
    annotation: ParamAnnotation | null
  }
  hint: {
    label: string             // UI에 표시할 라벨 (한국어)
    min?: number
    max?: number
    step?: number
    options?: string[]
    widget: "slider" | "number" | "text" | "toggle" | "select" | "range"
  }
}

type ParamAnnotation = {
  label: string
  min?: number; max?: number; step?: number
  options?: string[]
}

type State = {
  source: string                    // 원본 코드, 불변
  editables: Editable[]             // 한 번 추출 후 source가 바뀔 때만 재계산
  edits: Map<string, string>        // id -> 새 raw 토큰
  history: Array<Map<string, string>>  // 되돌리기용 스냅샷
}
```

렌더링:

```
function render(state):
  let out = ""
  let cursor = 0
  for ed in editables sorted by range[0]:
    out += source.slice(cursor, ed.range[0])
    out += edits.get(ed.id) ?? ed.raw
    cursor = ed.range[1]
  out += source.slice(cursor)
  return out
```

## 파서 사양 (라인 토크나이저)

### 전처리

- BOM 제거. CRLF → LF 정규화는 **하지 않음** (오프셋 보존을 위해).
- 라인 잇기: 라인 끝이 ` _`(공백 + 언더스코어)면 다음 라인과 결합한
  **논리 라인**을 만들되, 오프셋은 원본 기준 유지.

### 라인 분류 (대소문자 무시, 들여쓰기 후 첫 키워드 기준)

| 분류 | 패턴 |
|---|---|
| `subStart` | `Sub <이름>(...)` / `Function <이름>(...)` / `Private Sub` 등 |
| `subEnd` | `End Sub` / `End Function` |
| `decl` | `Dim`, `Const`, `Static`, `Public`, `Private` 변수 선언 |
| `if` | `If ... Then` (단일 라인 또는 블록) |
| `elseIf` | `ElseIf ... Then` |
| `else` | `Else` |
| `endIf` | `End If` |
| `for` | `For <var> = <expr> To <expr> [Step <expr>]` |
| `next` | `Next [<var>]` |
| `while` | `While ...` / `Do While ...` / `Do Until ...` |
| `wend` | `Wend` / `Loop` |
| `with` | `With ...` / `End With` |
| `comment` | `'...` 또는 `Rem ...` 시작 |
| `blank` | 공백만 |
| `stmt` | 그 외 일반 문 |

### 토큰 스캔 (라인 안에서)

상태 머신으로 좌→우 스캔. 산출:

- `string`: `"..."`. 내부 `""`는 이스케이프된 따옴표.
- `number`: `[+-]?\d+(\.\d+)?`, `&H[0-9A-Fa-f]+` (16진).
- `bool`: `True` / `False` (식별자가 아닌 위치).
- `ident`: `[A-Za-z_][A-Za-z0-9_]*` (한글 식별자도 허용 — 유니코드 확장).
- `op`: `=`, `<`, `>`, `<=`, `>=`, `<>`, `+`, `-`, `*`, `/`, `\`, `Mod`, `&`, `And`, `Or`, `Not`.
- `punct`: `(`, `)`, `,`, `.`, `:`.
- `comment`: `'`부터 줄 끝.

각 토큰은 `[start, end)` 오프셋 보유.

### 인식 못 하는 라인

스캔만 수행하고 분류는 `stmt`로. **추출 결정은 §추출 규칙에서
컨텍스트가 맞을 때만** 한다 — 모르는 구조에서 무작정 리터럴을
빼지 않는다.

## 추출 규칙 (어떤 리터럴을 위젯으로 노출하는가)

다음 컨텍스트에서만 리터럴을 추출. 그 외는 무시.

### R1. `Const` 우변
```vb
Const LAST_ROW = 100        '→ 추출 (number)
Const SHEET_NAME = "매출"   '→ 추출 (string)
```

### R2. `Dim ... = ...` 초기값
```vb
Dim threshold As Long: threshold = 50000   '→ 50000 추출
```

### R3. `For` 양 끝과 `Step`
```vb
For i = 2 To 100 Step 1
  '→ 2, 100, 1 모두 추출 (kind: forBound)
```
단, `Step 1`은 step=1이 기본이라 제외해도 무방. **MVP에서는 추출하되
hint.label = "증가폭"으로**.

### R4. `If <var> <op> <리터럴>` 의 우변
```vb
If 금액 > 50000 Then    '→ 50000 추출 (kind: ifRhs)
If name = "매출" Then   '→ "매출" 추출
```
좌변이 변수/함수호출, 우변이 리터럴인 경우만. 양변이 리터럴이면 추출
안 함(상수 비교 = 데드 코드 가능성).

### R5. 알려진 함수 호출의 리터럴 인자

화이트리스트:
- `Sheets("...")`, `Worksheets("...")`, `Workbooks("...")`
- `Range("...")`, `Cells(<n>, <n>)`, `.Range("...")`
- `MsgBox "..."`, `MsgBox "...", <n>`
- `InputBox("...")`

이 함수들의 인자 자리에 있는 문자열/숫자 리터럴만 추출.
사용자 정의 함수는 어노테이션이 없는 한 추출하지 않음 (안전).

### R6. 명시적 어노테이션 `'@param ...`

R1~R5의 휴리스틱을 무시하고 **그 라인의 첫 리터럴을** 어노테이션
라벨로 노출.

### 제외 규칙

- 숫자 0, 1, -1: 의미 없는 경우가 많음. 단 R1/R3/R6 컨텍스트에서는 유지.
- 배열 인덱스 `arr(0)` 의 0
- `Mod 2` 같은 산술 상수
- 문자열 길이 0인 `""`
- 같은 토큰이 동일 위치에서 반복되면 첫 번째만 — 이 부분은 §그룹화 참고

### 그룹화 (Phase 2)

같은 raw 값이 한 Sub 안에 여러 번 등장 → 하나의 위젯으로 묶고
"3곳에서 사용"이라 표시. 변경 시 모두 동기 수정.

## 라벨 결정 우선순위

```
1. annotation.label    (가장 신뢰)
2. leadingComment      (예: "' 마지막 행" → "마지막 행")
3. 식별자 기반 추측     (LAST_ROW → "마지막 행", threshold → "임계값")
4. 컨텍스트 기반 기본값 (For 끝값 → "반복 끝", Sheets 인자 → "시트 이름")
5. 최후: "값 (라인 N)"
```

식별자→한국어 사전은 작은 표로 시작 (last/end/max/min/threshold/
limit/sheet/row/col/path/file 등 ~30개).

## UX 와이어프레임

```
┌──────────────────────────────┬──────────────────────────────┐
│ 코드 (좌측, 60%)             │ 컨트롤 (우측, 40%)            │
│                              │                              │
│  Sub 데이터정리()             │ ▼ 데이터정리                  │
│    ' 마지막 행                │   마지막 행                   │
│   ▣Const LAST_ROW = 100      │   ──●──── [100]              │
│    Dim i As Long             │                              │
│    For i = 2 To▣LAST_ROW     │   시트 이름                   │
│      If Cells(i,3).Value > ▣ │   [매출        ]             │
│          50000 Then          │                              │
│        Sheets("▣매출")       │   금액 임계값                 │
│          .Range("A1").Value  │   ──────●── [50000]         │
│          = "확인"            │                              │
│      End If                  │ ▼ 변경 사항                   │
│    Next i                    │   • LAST_ROW: 100 → 200      │
│  End Sub                     │   [되돌리기] [모두 초기화]   │
│                              │                              │
│  [전체 복사]                 │                              │
└──────────────────────────────┴──────────────────────────────┘
```

- 좌측의 `▣` 표시는 추출된 토큰 하이라이트. 클릭하면 우측 위젯으로
  스크롤 + 강조.
- 우측 위젯 라벨 클릭 → 좌측 해당 토큰으로 스크롤 + 강조.
- 좌측 코드를 사용자가 직접 편집하면 300ms 디바운싱 후 재추출.
  편집 중 일시적으로 파싱 실패해도 마지막 성공 상태의 위젯 유지
  (붉은 인디케이터로 표시).

## 파일 구조 (Phase 1 끝 시점)

```
index.html          단일 페이지. 인라인 CSS + <script type="module">
src/
  parser.js         라인 분류 + 토큰 스캔
  extractor.js      Editable 추출 규칙 R1~R6
  labeler.js        라벨 결정 (식별자 사전 포함)
  state.js          State 객체 + edits/history
  render.js         오프셋 머지로 출력 문자열 생성
  ui.js             패널/위젯 렌더링, 이벤트
samples/
  sample1.bas       실제로 돌려볼 AI 생성 매크로 (사용자가 제공)
DESIGN.md
```

ESM 모듈, 빌드 도구 없음. 브라우저에서 `index.html` 더블클릭으로 동작.

## Phase별 작업 분해

### Phase 0 — 골격 (반나절)
- `index.html` 3분할 레이아웃, 좌측 textarea, 우측 placeholder
- 상태 객체와 빈 렌더링 파이프라인

### Phase 1 — MVP (이게 "쓸 만한" 첫 지점)
- 라인 토크나이저 (§파서 사양)
- 추출 규칙 R1, R2, R3, R5(`Sheets`/`Range`/`MsgBox`만)
- 라벨 우선순위 1~4
- 위젯: 숫자(input + slider), 문자열(input)
- 라이브 머지 출력, 전체 복사 버튼

### Phase 2 — 사용성
- 추출 R4 (`If` 우변), R6 (`@param`)
- Sub 트리 그룹화, 좌↔우 클릭 매핑
- 같은 값 그룹화
- 되돌리기/모두 초기화

### Phase 3 — 안전성/완성도
- 좌측 직접 편집 → 디바운싱 재추출
- 변경 사항 diff 뷰 (원본 vs 현재)
- 셀 주소(`"A1:B10"`) 미니 그리드 시각화

### Phase 4 (보류)
- LLM 코드 요약 — 사용자 키 입력 옵션
- 패턴 라이브러리 (자주 쓰는 매크로 템플릿)

## 미해결 / 추후 결정

1. **어노테이션 형식 확정**: `'@param 라벨 min=.. max=..` 로 갈 것인가,
   `'@param {label:"...", min:..}` JSON으로 갈 것인가. 전자가 사람
   친화적, 후자가 파싱 단순. → 일단 전자로.
2. **한글 인코딩**: 클립보드는 UTF-8로 들어오지만 일부 사용자가
   cp949 파일을 가질 수 있음. 파일 업로드 기능을 넣게 되면 인코딩
   감지 필요. MVP는 붙여넣기만 지원하므로 보류.
3. **저장**: 변경된 코드를 다시 작업 시 가져오려면 `localStorage`
   세션 복원. Phase 2 후반에 결정.
4. **샘플 매크로 부재**: 휴리스틱 튜닝을 위해 실제 AI 생성 매크로
   2~3개가 필요. 구현 들어가기 전 사용자에게 요청 예정.

## 성공 기준

비개발자 사용자에게 AI가 만든 100줄 내외의 VBA 매크로를 주고:
- 시트 이름, 행 범위, 임계값을 위젯만 만져 정확히 바꾸고
- 결과 코드를 Excel에 붙여넣어 정상 동작하기까지
- **소요 시간 1분 이내, 수동 코드 편집 0회**.
